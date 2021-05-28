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

-- Update cell map with values from CSV input.
updateRow :: Map String Int -> CellMap -> [String] -> CellMap
updateRow _ m [] = m
updateRow id_map m (l:f:id:_:_:vs) = case M.lookup id id_map of
                                      Just row -> updateCells  row 5 vs m
                                      otherwise -> m
updateRow _ m _ = m

updateCells :: Int -> Int -> [String] -> CellMap -> CellMap
updateCells _ _ [] m = m
updateCells rnum cnum (v:vs) m = updateCells rnum (cnum+1) vs $ updateCell rnum cnum v m
--updateCells _ _ _ m = m

updateCell :: Int -> Int -> String -> CellMap -> CellMap
updateCell rnum cnum v m = M.update (replaceVal v) (rnum, cnum) m

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
  , opt_values = "-"           &= typFile &= help "scores in CSV" &= name "values"
  , opt_input = "input.xlsx" &= typFile &= help "Input XLSX file name" &= name "input"
  , opt_output = "example.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  }
  &= summary "rupdate v0.1, (C) 2021 John Noll"
  &= program "main"


-- XXXjn this relies on there being a value at every valid cell in the input.  Seems to work.
projectCol :: Int -> CellMap -> [(Int, Int, String)]
projectCol target m = map toVal $ catMaybes $ map (\((r, c), cell) -> if c == target then Just (r, c, (_cellValue $ cell))  else Nothing) $ M.toAscList m

toVal :: (Int, Int, Maybe CellValue) -> (Int, Int, String)
toVal (r, c, Just (CellText v)) = (r, c, T.unpack v)
toVal (r, c, Just (CellDouble v)) = (r, c, show v)
toVal (r, c, Just (CellBool v)) = (r, c, show v)
toVal (r, c, Just (CellRich v)) = (r, c, "(rich text)")
toVal (r, c, Just (CellError v)) = (r, c, show v)
toVal (r, c, Nothing) = (r, c, "-9")

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
      id_map = M.fromList $ map (\(r, _, v) -> (v, r)) $ projectCol 2 cells
      cells' = foldl (updateRow id_map) cells vals
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