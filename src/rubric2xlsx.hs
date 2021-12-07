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
import Data.List (stripPrefix, foldl')
import Data.Map (Map)
import qualified Data.Map as M
import Data.Maybe (catMaybes)
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin)
import Text.CSV 
import Text.Printf (printf)
import Text.Regex.PCRE 
import Text.StringTemplate 


getCSV :: String ->  [[String]]
getCSV contents = case parseCSV "/dev/stderr" contents of
    Left s -> [["error", show $ s]]
    Right rows -> rows

colWidth :: Int -> Double -> ColumnsProperties
colWidth col width = ColumnsProperties { cpMin = col, cpMax = col, cpWidth = Just width, cpStyle = Nothing, cpHidden = False, cpCollapsed = False, cpBestFit = False }

colWidths :: [ColumnsProperties]
colWidths = [ colWidth 1  4.0 -- section
            , colWidth 2 18.0 -- question num
            , colWidth 3 65.0 -- question
            , colWidth 4  5.0 -- answer
            , colWidth 5 12.0 -- goto if no
            , colWidth 6 15.0 -- goto if yes
            , colWidth 7 45.0 -- comment
            ]

---------------------------------------------------------------------------------------------
-- Add formula to column 4 cells for setting value automatically
addFormulas :: [[String]] -> T.Text -> CellMap -> CellMap
addFormulas rows formula cmap = foldl' (addFormula formula) cmap $ zip [1..] rows

replaceFormula :: T.Text -> Cell -> Maybe Cell
replaceFormula f c = Just c { _cellFormula = Just $ CellFormula { _cellfAssignsToName = False, _cellfCalculate = False,  _cellfExpression = NormalFormula $ Formula { unFormula = f } } }

updateFormula :: Int -> Int -> T.Text -> CellMap -> CellMap
updateFormula rnum cnum formula cmap = M.update (replaceFormula formula) (rnum, cnum) cmap

addFormula :: T.Text -> CellMap -> (Int, [String]) -> CellMap
addFormula _ cmap (rnum, (h:r)) | rnum < 5 = cmap
addFormula _ cmap (rnum, (h:r)) | (h =~ rx "^Part") = cmap
addFormula _ cmap (rnum, (h:r)) | (h =~ rx "^Sect") = cmap
addFormula f cmap (rnum, (_:_:_:v:_)) = updateFormula rnum 4 f cmap
addFormula _ cmap (_, _) = cmap




---------------------------------------------------------------------------------------------
-- Conditional formatting, based on answer in column 4

dxfFillPattern :: T.Text -> Dxf
dxfFillPattern color = def {_dxfFill = Just $ Fill { _fillPattern = Just $ def { _fillPatternBgColor = Just $ def { _colorARGB = Just color } } } }

-- XXXjn it appears we are allowed only two Dxf styles (0 & 1)
yesFillPattern :: Dxf
yesFillPattern = dxfFillPattern "00FF00"

yesFillPatternNum :: Int
yesFillPatternNum = 0

noFillPattern :: Dxf
noFillPattern =  dxfFillPattern "FF6666"

noFillPatternNum :: Int
noFillPatternNum = 1

-- this does not register :(
alertFillPattern :: Dxf
alertFillPattern = dxfFillPattern "33FFFF"

alertFillPatternNum :: Int
alertFillPatternNum = 3


rubricStyleSheet :: StyleSheet 
rubricStyleSheet = minimalStyleSheet { _styleSheetDxfs = [ yesFillPattern, noFillPattern, alertFillPattern ] }
                                           
makeConditionals :: [[String]] -> [(SqRef, [CfRule])]
makeConditionals rows = catMaybes $ concat $ Prelude.map makeCondRow $ zip [1..] rows

createCondition :: Int -> Int -> T.Text -> Int -> (SqRef, [CfRule])
createCondition r c formula style = (SqRef [singleCellRef (r, c)], [CfRule { _cfrCondition = Expression $ Formula { unFormula = formula }, _cfrDxfId = Just style, _cfrPriority = 1, _cfrStopIfTrue = Just False } ])

highlightCond :: Int -> String -> T.Text
highlightCond rnum template = let tmpl = newSTMP template :: StringTemplate String
                              in T.pack $ render $ setManyAttrib [("row", rnum), ("parent", rnum-1), ("grandparent", rnum-2)] tmpl
cond :: String -> T.Text
cond val =  T.pack $ val

cond' :: Int -> String -> T.Text
cond' row val = cond $ printf "D%d=\"%s\"" row val

cond'' :: Int -> String -> T.Text
cond'' row val = cond$ printf "AND(D%d=\"%s\", D%d=\"?\")" (row-1) val row


makeCondRow :: (Int, [String]) -> [Maybe (SqRef, [CfRule]) ]
makeCondRow (rnum, (h:r)) | rnum < 5 = [Nothing]
makeCondRow (rnum, (h:r)) | (h =~ rx "^Part") = [Nothing]
makeCondRow (rnum, (h:r)) | (h =~ rx "^Sect") = [Nothing]
makeCondRow (rnum, (p:_:_:expr:n1:n2:_)) =
  [ Just $ createCondition rnum 5 (cond' rnum "No") 0
  , Just $ createCondition rnum 6 (cond' rnum "Yes") 0
  , Just $ createCondition rnum 4 (highlightCond rnum expr) 0
  ]
makeCondRow (_, _) = [Nothing]

---------------------------------------------------------------------------------------------
-- The base rubric formatting, with widths and colors.
formattedCells :: M.Map (Int, Int) FormattedCell -> Xlsx
formattedCells cells = formatWorkbook [("Sheet1", cells)] rubricStyleSheet

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
makeRow (rnum, (p:n:q:_:n1:n2:c:r)) = let bgColor = qColor n in -- this is the question row
                                       [ makeCell rnum 1 p def { fill = bgColor }
                                       , makeCell rnum 2 n def { fill = bgColor }
                                       , makeCell rnum 3 q def { fill = bgColor, indent = (qIndent n) }
                                       , makeCell rnum 4 ("?" :: String) def { fill = "FFFF66", border = lineBorder } -- reset value to '?'
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
      conds = makeConditionals rubric
--      formulaCells = addFormulas rubric "IF(\"Yes\", \"No\", \"Yes\")" $ _wsCells ws
      xlsx = fmt & atSheet "Sheet1" ?~ ws { _wsDataValidations = validations
                                          , _wsColumnsProperties = colWidths
                                          , _wsConditionalFormattings = M.fromList conds
--                                          , _wsCells = formulaCells
                                          }
      styleSheet = case parseStyleSheet $ _xlStyles xlsx of
        Left e -> minimalStyleSheet
        Right s -> s  { _styleSheetDxfs = [ yesFillPattern, noFillPattern ] } 
  L.writeFile (opt_output args) $ fromXlsx ct xlsx { _xlStyles = renderStyleSheet styleSheet}

