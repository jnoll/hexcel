{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE PatternGuards #-} -- for string prefix pattern
{-# LANGUAGE DeriveDataTypeable #-} -- for cmdargs
module Main where
import Codec.Xlsx
import Codec.Xlsx.Formatted (formatWorkbook, FormattedCell(..))
import Codec.Xlsx.SimpleFormatted
import Control.Lens
import qualified Data.Aeson.Key as Key
import qualified Data.Aeson.KeyMap as KeyMap
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Char8 as LU 
import Data.List (concat, intercalate, map, nub, foldl')
--import Data.Map (Map)
import qualified Data.HashMap.Strict as HM
import qualified Data.Map as M
import Data.Maybe (fromJust, fromMaybe, catMaybes, isJust)
import Data.Scientific (FPFormat(Generic), formatScientific)
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import Data.Vector as V hiding (concat, foldl', map, zip)
import qualified Data.Yaml as Y
import Data.YamlRubric (getYAML)
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.Exit (die, exitFailure, exitSuccess, ExitCode(..))
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin, openTempFileWithDefaultPermissions)
import System.Directory (renameFile)
import System.FilePath (takeExtension, takeBaseName, (</>), (<.>))
import Text.CSV
import Text.Pretty.Simple (pPrint, pPrintString)


toInt :: String -> Int
toInt s = read s


makeQmapEntry :: CellMap -> M.Map (String, String) Int -> Int -> M.Map (String, String) Int
makeQmapEntry cm qm r = let sec = M.lookup (r, 1) cm
                            q   = M.lookup (r, 2) cm
                        in if isJust sec && isJust q then
                             let (CellText sec') = fromMaybe (CellText "") $ (_cellValue $ fromJust sec) 
                                 (CellText q')   = fromMaybe (CellText "") $ (_cellValue $ fromJust q)
                             in M.insert (T.unpack sec', T.unpack q') r qm else qm

-- Build a "question map" (qmap) of (section, question number) to row number.
-- The numCell limit is a bit lame, but I don't know how to get the number of rows;
-- but it can't be more than the number of cells.
makeQmap :: CellMap -> M.Map (String, String) Int
makeQmap cmap = let numCell = M.size cmap in foldl' (makeQmapEntry cmap) M.empty [1..numCell]


-- Update cell map with answer to selected questions in col 4, and comment in col 7.
updateCell :: M.Map (String, String) Int -> CellMap -> [String] -> CellMap
updateCell _ m [] = m
updateCell _ m ("Sec":_) = m
updateCell _ m ("SEC":_) = m
updateCell _ m ("sec":_) = m
updateCell qmap m ("x":field:v:_) = case M.lookup ("x", field) qmap of
                                    Just row -> M.update (replaceVal "x") (row, 1) $ M.update (replaceVal v) (row, 3) m
                                    otherwise -> m
updateCell qmap m (sec:qnum:v:com:_) = case M.lookup (sec, qnum) qmap of
                                    Just row -> M.update (replaceVal com) (row, 7) $ M.update (replaceVal v) (row, 4) m
                                    otherwise -> m
updateCell _ m _ = m

replaceVal :: String -> Cell -> Maybe Cell
replaceVal v c = Just c { _cellValue = Just $ CellText $ T.pack v } 

--toEl :: (T.Text, Y.Value) -> [String]
toEl :: (Key.Key, Y.Value) -> [String]
-- create elements from Array.  The element tag is derived from singular from of surrounding tag, derived from Array key.
toEl (k, (Y.Array a)) =
    let --k' = T.unpack k
        k' = Key.toString k
        cs = intercalate ", " $ map (\v -> toVal $ Just v) $ V.toList a
    in ["x", k', cs, "comment"]
toEl (k, v) = let k' = Key.toString k 
              in ["x", k', toVal $ Just v, "comment"]

toVal :: Maybe Y.Value -> String
toVal (Just (Y.String s)) = T.unpack s
toVal (Just (Y.Number n)) = formatScientific Generic Nothing n
toVal (Just (Y.Bool True)) = "true"
toVal (Just (Y.Bool False)) = "false"
toVal (Just (Y.Null)) = ""
toVal Nothing = "nothing"
toVal (Just (Y.Array a)) = intercalate ", " $ map (\v -> toVal $ Just v) $ V.toList a
toVal _ = "foo unknown YAML element"

toVals :: Y.Value -> [[String]]
--toVals (Y.Object m) = map toEl $ (HM.toList m :: [(T.Text, Y.Value)])
toVals (Y.Object m) = map toEl $ KeyMap.toList m 
toVals (Y.Array a) = concat $ map toVals $ V.toList a
toVals _ = [["unknown YAML element"]]

getCSV :: String ->  [[String]]
getCSV contents = case parseCSV "/dev/stderr" contents of
    Left s -> [["error", show $ s]]
    Right rows -> rows

-- Make a string into a regular expression.  Only to make type clear.
rx :: String -> String
rx s = s::String


-- Main entry point

data Options = Options {
    opt_sheet :: T.Text
  , opt_values :: FilePath
  , opt_input :: FilePath
  , opt_output :: FilePath
  , opt_yaml :: Bool
  , opt_rubric :: Bool
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
    opt_sheet = "Sheet1" &= typ "String" &= help "Sheet name of input to update" &= name "sheet"
  , opt_values = "-" &= typFile &= help "List of (section, question, value) triples to set as answers to questions (CSV)" &= name "values"
  , opt_yaml = False &= typ "Bool" &= help "Is input YAML?" &= name "yaml"
  , opt_input = "input.xlsx" &= typFile &= help "Input XLSX file name" &= name "input"
  , opt_output = "example.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  , opt_rubric = False &= typ "Bool" &= help "Is input a rubric (in YAML format)?" &= name "rubric"
  }
  &= summary "rupdate v0.1, (C) 2021 John Noll"
  &= program "main"


-- XXXjn this relies on there being a value at every valid cell in the rubric.  Seems to work.
projectCol :: Int -> CellMap -> [String]
projectCol target m = map (\(CellText v) -> T.unpack v) $ Data.Maybe.catMaybes $ map (\((r, c), cell) -> if c == target then (_cellValue $ cell)  else Nothing) $ M.toAscList m


populate :: L.ByteString -> [[String]] -> T.Text -> Xlsx
populate excel vals sheetnm =
  let xlsx = toXlsx excel
      sheet =  fromJust (xlsx ^? ixSheet sheetnm)
      cells = _wsCells sheet


      qmap = makeQmap cells
      cells' = foldl' (updateCell qmap) cells vals
  in
    xlsx & atSheet sheetnm ?~ sheet { _wsCells = cells' }

  
  
main :: IO ()
main = do
  args <- cmdArgs defaultOptions
  ct <- getPOSIXTime
  let ext = takeExtension (opt_values args)
      isYaml = if (opt_yaml args) == False && (opt_values args) /= "-" && (ext == ".yml" || ext == ".yaml" || ext == ".YML" || ext == ".YAML")
             then True else (opt_yaml args)
  -- putStrLn $  (show yaml)
  -- -v - means read from stdin.
  valh <- if (opt_values args) == "-" then return stdin else openFile (opt_values args) ReadMode
  valc <- hGetContents valh
  excel <- L.readFile (opt_input args)


  -- It's considered bad practice to read from and write to the name file, so
  -- write to a temp file, then rename it.
  (tmpNm, tmph) <- openTempFileWithDefaultPermissions "." "output.xlsx" -- becomes "outputNNN.xlsx"

  if isYaml then
      case Y.decodeEither' $ LU.pack valc of
                     Left err -> do die $ Y.prettyPrintParseException err
                     Right y -> if (opt_rubric args) == True then do
                                  putStrLn $ printCSV $ getYAML y
                                  L.hPut tmph $ fromXlsx ct $ populate excel (getYAML y) (opt_sheet args)
                                else
                                  L.hPut tmph $ fromXlsx ct $ populate excel (toVals y) (opt_sheet args)
    else
    let vals = getCSV valc in do
      L.hPut tmph $ fromXlsx ct $ populate excel vals (opt_sheet args)

  hClose tmph

  renameFile tmpNm  (opt_output args)
  exitSuccess
