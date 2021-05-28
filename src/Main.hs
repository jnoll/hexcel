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
import Data.Maybe (fromJust)
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin, openTempFileWithDefaultPermissions)
import System.Directory (renameFile)
import Text.CSV 


toInt :: String -> Int
toInt s = read s

-- Update cell map  from a CSV version of the worksheet.
-- OK for small worksheets, but if your working with Big Data you'll need a more
-- efficient approach that probably creates a sorted list of cells.
updateRows :: Int -> [[String]]  -> CellMap -> CellMap
updateRows _ []  m = m
updateRows rnum (r:rs) m = updateRows (rnum + 1) rs $ updateRow rnum 1 r m

updateRow :: Int -> Int -> [String] -> CellMap -> CellMap
updateRow _ _ [] m = m
updateRow rnum cnum (c:cs) m = updateRow rnum (cnum+1) cs $ updateCell' rnum cnum c m

updateCell' :: Int -> Int -> String -> CellMap -> CellMap
updateCell' rnum cnum v m = M.update (replaceVal v) (rnum, cnum) m

-- Update cell map with new value at (row, col)
updateCell :: CellMap -> [String] -> CellMap
updateCell m ("row":_) = m
updateCell m ("Row":_) = m
updateCell m ("ROW":_) = m
updateCell m (row:col:val:_) = M.update (replaceVal val) (toInt row, toInt col) m
updateCell m _ = m

replaceVal :: String -> Cell -> Maybe Cell
replaceVal v c = Just c { _cellValue = Just $ CellText $ T.pack v } 

getCSV :: String ->  [[String]]
getCSV contents = case parseCSV "/dev/stderr" contents of
    Left s -> [["error", show $ s]]
    Right rows -> rows



-- Main entry point

data Options = Options {
  opt_values :: FilePath
  , opt_sheet :: T.Text
  , opt_input :: FilePath
  , opt_output :: FilePath
  , opt_update :: Bool
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
  opt_values = "-" &= typFile &= help "List of (row, column, value) triples to set (CSV)" &= name "values"
  , opt_sheet = "Sheet1" &= typ "String" &= help "Sheet name to update" &= name "sheet"
  , opt_input = "rubric.xlsx" &= typFile &= help "Input XLSX file name" &= name "input"
  , opt_output = "example.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  , opt_update = False &= typ "Boolean" &= help "Update worksheet cell values from a CSV version of the worksheet" &= name "update"
  }
  &= summary "hexcel v0.1, (C) 2021 John Noll"
  &= program "main"


main :: IO ()
main = do
  args <- cmdArgs defaultOptions
  ct <- getPOSIXTime
  -- "-v -" means read values from stdin.  Useful for values provided by a pipeline.
  valh <- if (opt_values args) == "-" then return stdin else openFile (opt_values args) ReadMode
  valc <- hGetContents valh

  excel <- L.readFile (opt_input args)

  let xlsx = toXlsx excel
      sheet =  fromJust (xlsx ^? ixSheet (opt_sheet args) )
      vals = getCSV valc
      cells = if opt_update args then
                updateRows 1 vals (_wsCells $ sheet) 
              else
                foldl updateCell (_wsCells $ sheet) vals
      xlsx' = xlsx & atSheet (opt_sheet args) ?~ sheet { _wsCells = cells }

  -- It's considered bad practice to read from and write to the name file, so
  -- write to a temp file, then rename it.
  (tmpNm, tmph) <- openTempFileWithDefaultPermissions "." "output.xlsx" -- becomes "outputNNN.xlsx"
  L.hPut tmph $ fromXlsx ct xlsx'
  hClose tmph

  renameFile tmpNm  (opt_output args)
