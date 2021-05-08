{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE PatternGuards #-} -- for string prefix pattern
module Main where
import Codec.Xlsx
import Codec.Xlsx.Formatted
import Control.Lens
import qualified Data.ByteString.Lazy as L
import Data.List (stripPrefix)
import Data.Map (Map)
import qualified Data.Map as M
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import Text.CSV 
import Text.Regex.PCRE 

main :: IO ()
main = do
  contents <- getContents
  ct <- getPOSIXTime

  let
      csv = getCSV contents
      fmt = formattedCells $ makeCells csv
      (n, ws) = head $ _xlSheets fmt
      validations = addMenu csv
      xlsx = fmt & atSheet "Sheet1" ?~ ws { _wsDataValidations = validations, _wsColumnsProperties = colWidths }
  L.writeFile "example.xlsx" $ fromXlsx ct xlsx

colWidth :: Int -> Double -> ColumnsProperties
colWidth col width = ColumnsProperties { cpMin = col, cpMax = col, cpWidth = Just width, cpStyle = Nothing, cpHidden = False, cpCollapsed = False, cpBestFit = False }

colWidths :: [ColumnsProperties]
colWidths = [ colWidth 1  4.0 -- section
            , colWidth 2  9.0 -- question num
            , colWidth 3 65.0 -- question
            , colWidth 4  5.0 -- answer
            , colWidth 5 12.0 -- goto if no
            , colWidth 6 15.0 -- goto if yes
            , colWidth 7 45.0 -- comment
            ]

getCSV :: String ->  [[String]]
getCSV contents = case parseCSV "/dev/stderr" contents of
    Left s -> [["error", show $ s]]
    Right rows -> rows

formattedCells :: [((Int, Int), FormattedCell)] -> Xlsx
formattedCells cells = formatWorkbook [("Sheet1", M.fromList cells)] minimalStyleSheet


makeCells :: [[String]] -> [((Int, Int), FormattedCell)]
makeCells rows = concat $ Prelude.map makeRow $ zip [1..] rows

rx :: String -> String
rx s = s::String

makeRow :: (Int, [String]) -> [ ((Int, Int), FormattedCell) ]
makeRow (rnum, (h:r)) | rnum < 5 = Prelude.map (makeCell rnum "FFFFFF" 0 noBorder) $ zip [1..] (h:r)
makeRow (rnum, (h:r)) | (h =~ rx "^Part") = Prelude.map (makeCell rnum "99CCFF" 0 noBorder) $ zip [1..] (h:r)
makeRow (rnum, (h:r)) | (h =~ rx "^Sect") = Prelude.map (makeCell rnum "B2FF66" 0 noBorder) $ zip [1..] (h:r)
makeRow (rnum, (p:n:q:a:n1:n2:c:r))                 = let color = qColor n in
                                             [ makeCell rnum color 0 noBorder (1, p)
                                             , makeCell rnum color 0 noBorder (2, n)
                                             , makeCell rnum color (qIndent n) noBorder (3, q)
                                             , makeCell rnum "FFFF66" 0 lineBorder (4, a)
                                             , makeCell rnum color 0 noBorder (5, n1)
                                             , makeCell rnum color 0 noBorder (6, n2)
                                             , makeCell rnum color 0 lineBorder (7, c)
                                             ] ++ (Prelude.map (makeCell rnum color 0 noBorder) $ zip [8..] r)
makeRow (rnum, r)                         = Prelude.map (makeCell rnum "FFFFFF" 0 noBorder) $ zip [1..] (r)

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

-- question indent. Regexp rather than recursion seems like a hack.
qIndent :: String -> Int
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9][.][0-9][.][0-9]") = 5;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9][.][0-9]") = 4;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9][.][0-9]") = 3;
qIndent q | (q =~ rx "^[0-9][.][0-9][.][0-9]") = 2;
qIndent q | (q =~ rx "^[0-9][.][0-9]") = 1;
qIndent _  = 0;

noBorder :: Border
noBorder =  def { _borderLeft = noBorderStyle
                    , _borderRight = noBorderStyle
                    , _borderTop = noBorderStyle
                    , _borderBottom = noBorderStyle
                    }
lineBorder :: Border
lineBorder =  def { _borderLeft = lineBorderStyle
                      , _borderRight = lineBorderStyle
                      , _borderTop = lineBorderStyle
                      , _borderBottom = lineBorderStyle
                      }

makeCell :: Int -> T.Text -> Int -> Border -> (Int, String) -> ((Int, Int), FormattedCell)
--makeCell rnum fill indent border (cnum, val) | rnum < 6 = ((rnum, cnum), formatCell val fill False (Just border) indent)
--makeCell rnum fill indent (cnum, val) | cnum == 4 = ((rnum, cnum), formatCell val fill False lineBorder indent)
--makeCell rnum fill indent border (cnum, val) | cnum == 7 = ((rnum, cnum), formatCell val fill False (Just border) indent)
makeCell rnum fill indent border (cnum, val)  = ((rnum, cnum), formatCell val fill False (Just border) indent)


fillColor :: T.Text -> Maybe Fill
fillColor c = Just def { _fillPattern = Just def { _fillPatternBgColor = Just def { _colorARGB = Just c }
                                                 , _fillPatternFgColor = Just def { _colorARGB = Just c }
                                                 , _fillPatternType = Just PatternTypeSolid } }

formatCell val fill bold border indent = def { _formattedCell = def { _cellValue = Just $ CellText $ T.pack val }
                                             , _formattedFormat = def { _formatFill = fillColor fill 
                                                                      , _formatFont = Just def { _fontBold = Just bold, _fontName = Just "Calibri" }
                                                                      , _formatBorder = border
                                                                      , _formatAlignment = Just def { _alignmentIndent = Just indent
                                                                                                    , _alignmentHorizontal = Just CellHorizontalAlignmentLeft
                                                                                                    }
                                                                      }
                                             }



noBorderStyle :: Maybe BorderStyle
noBorderStyle = Just $ def { _borderStyleLine = Just LineStyleNone, _borderStyleColor = Just def { _colorARGB = Just "000000" } }

lineBorderStyle :: Maybe BorderStyle
lineBorderStyle = Just $ def { _borderStyleLine = Just LineStyleThin, _borderStyleColor = Just def { _colorARGB = Just "000000" } }


addMenu :: [[String]] -> Map SqRef  DataValidation
addMenu rows = M.fromList $ map addMenu' $ zip [1..] rows

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
