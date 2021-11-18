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
import Data.List (stripPrefix)
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (catMaybes)
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin)
import Text.CSV 
import Text.Regex.PCRE 



getCSV :: String ->  [[String]]
getCSV contents = case parseCSV "/dev/stderr" contents of
    Left s -> [["error", show $ s]]
    Right rows -> rows

colWidth :: Int -> Double -> ColumnsProperties
colWidth col width = ColumnsProperties { cpMin = col, cpMax = col, cpWidth = Just width, cpStyle = Nothing, cpHidden = False, cpCollapsed = False, cpBestFit = False }

colWidths :: [ColumnsProperties]
colWidths = [ colWidth 1  4.0 -- section
            , colWidth 2 15.0 -- question num
            , colWidth 3 65.0 -- question
            , colWidth 4  5.0 -- answer
            , colWidth 5 12.0 -- goto if no
            , colWidth 6 15.0 -- goto if yes
            , colWidth 7 45.0 -- comment
            ]


formattedCells :: M.Map (Int, Int) FormattedCell -> Xlsx
formattedCells cells = formatWorkbook [("Sheet1", cells)] minimalStyleSheet


makeCells :: [[String]] -> [((Int, Int), FormattedCell)]
makeCells rows = concat $ Prelude.map makeRow $ zip [1..] rows

-- Make a string into a regular expression.  Only to make type clear.
rx :: String -> String
rx s = s::String

-- The rubric has a plain header in the top four rows, a line of column headings in row 5,
-- then a few sections comprising a section header and one or more questions.
makeRow :: (Int, [String]) -> [ ((Int, Int), FormattedCell) ]
makeRow (rnum, (h:r)) | rnum < 5 = Prelude.map (\(c, v) -> makeCell rnum c v def) $ zip [1..] (h:r)
makeRow (rnum, (h:r)) | (h =~ rx "^Part") = Prelude.map (\(c, v) -> makeCell rnum c v def { fill = "99CCFF"}) $ zip [1..] (h:r)
makeRow (rnum, (h:r)) | (h =~ rx "^Sect") = Prelude.map (\(c, v) -> makeCell rnum c v def { fill = "B2FF66"}) $ zip [1..] (h:r)
makeRow (rnum, ("x":n:r)) = let bgColor = "FFDDFF"
                                color = "0000FF"
                            in -- this is the row containing submission data
                                          [ makeCell rnum 1 ("x" :: String) def { fill = bgColor, color = color, font = "Courier" }
                                          , makeCell rnum 2 n def { fill = bgColor, color = color, font = "Courier" }
                                          , makeCell rnum 3 ("" :: String) def { fill = bgColor, color = color, font = "Courier" }
                                          , makeCell rnum 4 ("" :: String) def { fill = bgColor }
                                          , makeCell rnum 5 ("" :: String) def { fill = bgColor }
                                          , makeCell rnum 6 ("" :: String) def { fill = bgColor }
                                          , makeCell rnum 7 ("" :: String) def { fill = bgColor, border = lineBorder }
                                          ] ++ (Prelude.map (\(c, v) -> makeCell rnum c v def { fill = bgColor }) $ zip [8..] r)
makeRow (rnum, (p:n:q:a:n1:n2:c:r)) = let bgColor = qColor n in -- this is the question row
                                       [ makeCell rnum 1 p def { fill = bgColor }
                                       , makeCell rnum 2 n def { fill = bgColor }
                                       , makeCell rnum 3 q def { fill = bgColor, indent = (qIndent n) }
                                       , makeCell rnum 4 a def { fill = "FFFF66", border = lineBorder }
                                       , makeCell rnum 5 n1 def { fill = bgColor }
                                       , makeCell rnum 6 n2 def { fill = bgColor }
                                       , makeCell rnum 7 c def { fill = bgColor, border = lineBorder }
                                       ] ++ (Prelude.map (\(c, v) -> makeCell rnum c v def { fill = bgColor }) $ zip [8..] r)
makeRow (rnum, r) = Prelude.map (\(c, v) -> makeCell rnum c v def { fill = "FFFFFF" }) $ zip [1..] (r)

-- question color, used to provide background colors for different sections.
qColor :: String -> T.Text
qColor q | (q =~ rx "^(2|4|6|8)[.]1") = "FFCC99"
qColor q | (q =~ rx "^(2|4|6|8)[.]2") = "FFFFCC"
qColor q | (q =~ rx "^(2|4|6|8)[.]3") = "FFCCCC"
qColor q | (q =~ rx "^(2|4|6|8)")     = "FFCC99"
qColor q | (q =~ rx "^(1|3|5|7)[.]1") = "F5F5F5"
qColor q | (q =~ rx "^(1|3|5|7)[.]2") = "FF99CC"
qColor q | (q =~ rx "^(1|3|5|7)[.]3") = "E5CCFF"
qColor q | (q =~ rx "^(1|3|5|7)")     = "F5F5F5"
qColor q                              = "FFFFFF"

-- question indent. Regexp rather than recursion is a hack, but works OK for CSV input.
qIndent :: String -> Int
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9][.][0-9][.][0-9][.][0-9][.][0-9]") = 7;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9][.][0-9][.][0-9][.][0-9]") = 6;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9][.][0-9][.][0-9]") = 5;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9][.][0-9]") = 4;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9]") = 3;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9]") = 2;
qIndent q | (q =~ rx "^[0-9][.][0-9]") = 1;
qIndent _  = 0;


-- Column four of question lines gets a dropdown menu.
-- Dropdowns are part of data validations.
addMenu :: [[String]] -> Map SqRef  DataValidation
addMenu rows = M.fromList $ map addMenu' $ zip [1..] rows

-- regex to match question lines, which get menus.
qregex = "^[ \t]*[A-Z][ \t]*$" :: String -- Single letter indicates a question.

-- Anything that doesn't match one of the prefixes gets a menu.
addMenu' :: (Int, [String]) -> (SqRef, DataValidation)
addMenu' (n, h:_) | (h =~ qregex) = (SqRef [singleCellRef (n, 4)], defMenu) -- these get dropdown
addMenu' (n, _) = (SqRef [singleCellRef (n, 4)], def) -- these get defaul

defMenu :: DataValidation
defMenu = def { _dvAllowBlank       = False
              , _dvError            = Just "Incorrect data"
              , _dvErrorStyle       = ErrorStyleInformation
              , _dvErrorTitle       = Just "Click the down arrow to right of this cell to see your options"
              , _dvPrompt           = Just "Please use dropdown options only"
              , _dvPromptTitle      = Just "Score"
              , _dvShowDropDown     = False -- this has the opposite effect to what I expect
              , _dvShowErrorMessage = True
              , _dvShowInputMessage = True
              , _dvValidationType   = ValidationTypeList ["N/A","Yes","No"]
              }


-- Main entry point

data Options = Options {
  opt_rubric :: FilePath
  , opt_output :: FilePath
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
  opt_rubric = "rubric.csv"    &= typFile &= help "Rubric, in CSV format" &= name "rubric"
  , opt_output = "rubric.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  }
  &= summary "hexcel v0.1, (C) 2021 John Noll"
  &= program "main"


main :: IO ()
main = do
  args <- cmdArgs defaultOptions
  contents <- readFile (opt_rubric args) 

  ct <- getPOSIXTime

  let
      rubric = getCSV contents
      cells = M.fromList $ makeCells rubric
      fmt = formattedCells cells
      (n, ws) = head $ _xlSheets fmt
      validations = addMenu rubric
      xlsx = fmt & atSheet "Sheet1" ?~ ws { _wsDataValidations = validations, _wsColumnsProperties = colWidths }
  L.writeFile (opt_output args) $ fromXlsx ct xlsx

