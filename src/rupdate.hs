{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE PatternGuards #-} -- for string prefix pattern
{-# LANGUAGE DeriveDataTypeable #-} -- for cmdargs
module Main where
import Codec.Xlsx
import Codec.Xlsx.Formatted (formatWorkbook, FormattedCell(..))
import Codec.Xlsx.SimpleFormatted
import Control.Lens
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Char8 as LU 
import Data.List (intercalate, map, nub)
--import Data.Map (Map)
import qualified Data.HashMap.Strict as HM
import qualified Data.Map as M
import Data.Maybe (fromJust, fromMaybe, catMaybes, isJust)
import Data.Scientific (FPFormat(Generic), formatScientific)
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import Data.Vector as V hiding (foldl, map, zip)
import qualified Data.Yaml as Y
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.Exit (die, exitFailure, exitSuccess, ExitCode(..))
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin, openTempFileWithDefaultPermissions)
import System.Directory (renameFile)
import System.FilePath (takeExtension, takeBaseName, (</>), (<.>))
import Text.CSV 


toInt :: String -> Int
toInt s = read s


makeQmapEntry :: CellMap -> M.Map (String, String) Int -> Int -> M.Map (String, String) Int
makeQmapEntry cm qm r = let sec = M.lookup (r, 1) cm
                            q   = M.lookup (r, 2) cm
                        in if isJust sec && isJust q then
                             let (CellText sec') = fromMaybe (CellText "") $ (_cellValue $ fromJust sec) 
                                 (CellText q')   = fromMaybe (CellText "") $ (_cellValue $ fromJust q)
                             in M.insert (T.unpack sec', T.unpack q') r qm else qm



-- Update cell map with answer to selected questions in col 4, and comment in col 7.
updateCell :: M.Map (String, String) Int -> CellMap -> [String] -> CellMap
updateCell _ m [] = m
updateCell _ m ("Sec":_) = m
updateCell _ m ("SEC":_) = m
updateCell _ m ("sec":_) = m
updateCell qmap m ("x":field:v:_) = case M.lookup ("x", field) qmap of
                                    Just row -> M.update (replaceVal "") (row, 1) $ M.update (replaceVal v) (row, 3) m
                                    otherwise -> m
updateCell qmap m (sec:qnum:v:com:_) = case M.lookup (sec, qnum) qmap of
                                    Just row -> M.update (replaceVal com) (row, 7) $ M.update (replaceVal v) (row, 4) m
                                    otherwise -> m
updateCell _ m _ = m

replaceVal :: String -> Cell -> Maybe Cell
replaceVal v c = Just c { _cellValue = Just $ CellText $ T.pack v } 

--toEl :: (T.Text, Y.Value) -> [String]
toEl :: (T.Text, Y.Value) -> [String]
-- create elements from Array.  The element tag is derived from singular from of surrounding tag, derived from Array key.
toEl (k, (Y.Array a)) =
    let k' = T.unpack k
        cs = intercalate ", " $ map (\v -> toVal $ Just v) $ V.toList a
    in ["x", k', cs, "comment"]
toEl (k, v) = let k' = T.unpack k 
              in ["x", k', toVal $ Just v, "comment"]

toVal :: Maybe Y.Value -> String
toVal (Just (Y.String s)) = T.unpack s
toVal (Just (Y.Number n)) = formatScientific Generic Nothing n
toVal (Just (Y.Bool True)) = "true"
toVal (Just (Y.Bool False)) = "false"
toVal (Just (Y.Null)) = ""
toVal Nothing = "nothing"
toVal (Just (Y.Array a)) = intercalate ", " $ map (\v -> toVal $ Just v) $ V.toList a
toVal _ = "unknown YAML element"

toVals :: Y.Value -> [[String]]
toVals (Y.Object m) = map toEl $ (HM.toList m :: [(T.Text, Y.Value)])
toVals _ = [["unknown YAML element"]]

getYAML :: String -> [[String]]
getYAML contents = case Y.decodeEither' $ LU.pack contents of
                     Left err -> [[Y.prettyPrintParseException err]]
                     Right y -> toVals y
  
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
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
    opt_sheet = "Sheet1" &= typ "String" &= help "Sheet name of input to update" &= name "sheet"
  , opt_values = "-" &= typFile &= help "List of (section, question, value) triples to set as answers to questions (CSV)" &= name "values"
  , opt_yaml = False &= typ "Bool" &= help "Is input YAML?" &= name "yaml"
  , opt_input = "input.xlsx" &= typFile &= help "Input XLSX file name" &= name "input"
  , opt_output = "example.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  }
  &= summary "rupdate v0.1, (C) 2021 John Noll"
  &= program "main"


-- XXXjn this relies on there being a value at every valid cell in the rubric.  Seems to work.
projectCol :: Int -> CellMap -> [String]
projectCol target m = map (\(CellText v) -> T.unpack v) $ catMaybes $ map (\((r, c), cell) -> if c == target then (_cellValue $ cell)  else Nothing) $ M.toAscList m

populate :: L.ByteString -> [[String]] -> T.Text -> Xlsx
populate excel vals sheetnm =
  let xlsx = toXlsx excel
      sheet =  fromJust (xlsx ^? ixSheet sheetnm)
      cells = _wsCells sheet

      -- Build a map of (section, question number) to row number.
      -- This is more efficient, but at the cost of understandabiltity
      -- It ends up being 10N, which is silly for rubrics with 30-40 rows, but maybe OK
      -- for large rubrics like the final report that has 81 rows, or nearly 700 cells.
      qmap = M.fromList $ zip (zip (projectCol 1 cells) (projectCol 2 cells)) [1..]
      cells' = foldl (updateCell qmap) cells vals
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
                     Right y -> L.hPut tmph $ fromXlsx ct $ populate excel (toVals y) (opt_sheet args)
    else
    let vals = getCSV valc in do
      L.hPut tmph $ fromXlsx ct $ populate excel vals (opt_sheet args)

  hClose tmph

  renameFile tmpNm  (opt_output args)
  exitSuccess
