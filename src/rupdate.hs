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
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (fromJust, catMaybes)
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin, openTempFileWithDefaultPermissions)
import System.Directory (renameFile)
import Text.CSV 


toInt :: String -> Int
toInt s = read s


makeQmapEntry :: (Int, [String]) -> Maybe ((String, String), Int)
makeQmapEntry (r, sec:qnum:_) = Just ((sec, qnum), r)
makeQmapEntry (r, _) = Nothing

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
  opt_rubric :: FilePath
  , opt_sheet :: T.Text
  , opt_values :: FilePath
  , opt_input :: FilePath
  , opt_output :: FilePath
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
  opt_rubric = "rubric.xlsx"    &= typFile &= help "Rubric, in CSV format" &= name "rubric"
  , opt_sheet = "Sheet1" &= typ "String" &= help "Sheet name of input to update" &= name "sheet"
  , opt_values = "-"           &= typFile &= help "List of (section, question, value) triples to set as answers to questions (CSV)" &= name "values"
  , opt_input = "input.xlsx" &= typFile &= help "Input XLSX file name" &= name "input"
  , opt_output = "example.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  }
  &= summary "rupdate v0.1, (C) 2021 John Noll"
  &= program "main"


main :: IO ()
main = do
  putStrLn "I'm here, jim!"
  args <- cmdArgs defaultOptions
  ct <- getPOSIXTime
  -- -v - means read from stdin.
  valh <- if (opt_values args) == "-" then return stdin else openFile (opt_values args) ReadMode
  valc <- hGetContents valh
  rubricc <- readFile (opt_rubric args)

  excel <- L.readFile (opt_input args)

  let xlsx = toXlsx excel
      sheet =  fromJust (xlsx ^? ixSheet (opt_sheet args) )
      vals = getCSV valc
      rubric = getCSV rubricc
      qmap = M.fromList $ catMaybes $ map makeQmapEntry $ zip [1..] rubric
      cells = foldl (updateCell qmap) (_wsCells $ sheet) vals
      xlsx' = xlsx & atSheet (opt_sheet args) ?~ sheet { _wsCells = cells }

  -- It's considered bad practice to read from and write to the name file, so
  -- write to a temp file, then rename it.
  (tmpNm, tmph) <- openTempFileWithDefaultPermissions "." "output.xlsx" -- becomes "outputNNN.xlsx"
  L.hPut tmph $ fromXlsx ct xlsx'
  hClose tmph

  renameFile tmpNm  (opt_output args)
