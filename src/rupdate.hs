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
import Data.List (map, nub)
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (fromJust, fromMaybe, catMaybes, isJust)
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin, openTempFileWithDefaultPermissions)
import System.Directory (renameFile)
import Text.CSV 


toInt :: String -> Int
toInt s = read s


makeQmapEntry :: CellMap -> Map (String, String) Int -> Int -> Map (String, String) Int
makeQmapEntry cm qm r = let sec = M.lookup (r, 1) cm
                            q   = M.lookup (r, 2) cm
                        in if isJust sec && isJust q then
                             let (CellText sec') = fromMaybe (CellText "") $ (_cellValue $ fromJust sec) 
                                 (CellText q')   = fromMaybe (CellText "") $ (_cellValue $ fromJust q)
                             in M.insert (T.unpack sec', T.unpack q') r qm else qm



-- Update cell map with answer to selected questions in col 4, and comment in col 7.
updateCell :: Map (String, String) Int -> CellMap -> [String] -> CellMap
updateCell _ m [] = m
updateCell _ m ("Sec":_) = m
updateCell _ m ("SEC":_) = m
updateCell _ m ("sec":_) = m
updateCell qmap m (sec:qnum:v:com:_) = case M.lookup (sec, qnum) qmap of
                                    Just row -> M.update (replaceVal com) (row, 7) $ M.update (replaceVal v) (row, 4) m
                                    otherwise -> m
updateCell _ m _ = m

replaceVal :: String -> Cell -> Maybe Cell
replaceVal v c = Just c { _cellValue = Just $ CellText $ T.pack v } 

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
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
    opt_sheet = "Sheet1" &= typ "String" &= help "Sheet name of input to update" &= name "sheet"
  , opt_values = "-"           &= typFile &= help "List of (section, question, value) triples to set as answers to questions (CSV)" &= name "values"
  , opt_input = "input.xlsx" &= typFile &= help "Input XLSX file name" &= name "input"
  , opt_output = "example.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  }
  &= summary "rupdate v0.1, (C) 2021 John Noll"
  &= program "main"


-- XXXjn this relies on there being a value at every valid cell in the rubric.  Seems to work.
projectCol :: Int -> CellMap -> [String]
projectCol target m = map (\(CellText v) -> T.unpack v) $ catMaybes $ map (\((r, c), cell) -> if c == target then (_cellValue $ cell)  else Nothing) $ M.toAscList m

main :: IO ()
main = do
  args <- cmdArgs defaultOptions
  ct <- getPOSIXTime
  -- -v - means read from stdin.
  valh <- if (opt_values args) == "-" then return stdin else openFile (opt_values args) ReadMode
  valc <- hGetContents valh

  excel <- L.readFile (opt_input args)

  let xlsx = toXlsx excel
      sheet =  fromJust (xlsx ^? ixSheet (opt_sheet args) )
      vals = getCSV valc
      cells = _wsCells sheet
      -- This is more efficient, but at the cost of understandabiltity
      -- It ends up being 10N, which is silly for rubrics with 30-40 rows, but maybe OK
      -- for large rubrics like the final report that has 81 rows, or nearly 700 cells.
      qmap = M.fromList $ zip (zip (projectCol 1 cells) (projectCol 2 cells)) [1..]
      cells' = foldl (updateCell qmap) cells vals
      xlsx' = xlsx & atSheet (opt_sheet args) ?~ sheet { _wsCells = cells' }

  -- It's considered bad practice to read from and write to the name file, so
  -- write to a temp file, then rename it.
  (tmpNm, tmph) <- openTempFileWithDefaultPermissions "." "output.xlsx" -- becomes "outputNNN.xlsx"
  L.hPut tmph $ fromXlsx ct xlsx'
  hClose tmph

  renameFile tmpNm  (opt_output args)
--  putStrLn $ show $ zip (zip  (map (\(CellText v) -> v) $ catMaybes $ map (\((r, c), cell) -> if c == 1 then (_cellValue $ cell)  else Nothing) $ M.toAscList cells)
--                              (map (\(CellText v) -> v) $ catMaybes $ map (\((r, c), cell) -> if c == 2 then (_cellValue $ cell)  else Nothing) $ M.toAscList cells))
--                        [1..]
