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
--import Text.Regex.Posix

main :: IO ()
main = do
  contents <- getContents
  ct <- getPOSIXTime

  let
--      sheet = def & cellValueAt (1,2) ?~ CellDouble 42.0
--                  & cellValueAt (3,2) ?~ CellText "foo"
      csv = getCSV contents
--      sheet = makeSheet' def csv -- [["foo", "bar", "fro", "baz"], ["alpha", "beta", "gamma", "delta"]]
      fmt = formattedCells $ makeCells csv
      (n, ws) = head $ _xlSheets fmt
--      xlsx = def & atSheet "Sheet1" ?~ sheet { _wsDataValidations = validations, _wsCells = formattedCellMap fmt }
      validations = addMenu csv
      xlsx = fmt & atSheet "Sheet1" ?~ ws { _wsDataValidations = validations }
  L.writeFile "example.xlsx" $ fromXlsx ct xlsx

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
makeRow (rnum, (h:r)) | (h =~ rx "^Part") = Prelude.map (makeCell rnum partFill) $ zip [1..] (h:r)
makeRow (rnum, (h:r)) | (h =~ rx "^Sect") = Prelude.map (makeCell rnum sectFill) $ zip [1..] (h:r)
makeRow (rnum, (h:q:r)) | (q =~ rx "^(2|4|6)") = Prelude.map (makeCell rnum evenFill) $ zip [1..] (h:q:r)
makeRow (rnum, vals) = Prelude.map (makeCell rnum Nothing) $ zip [1..] (vals)

makeCell :: Int -> Maybe Fill -> (Int, String) -> ((Int, Int), FormattedCell)
makeCell rnum fill (cnum, val) = ((rnum, cnum), formatCell val fill False)

partFill :: Maybe Fill
partFill = Just def { _fillPattern = Just def { _fillPatternBgColor = Just def { _colorARGB = Just "CCFFF" }
                                              , _fillPatternFgColor = Just def { _colorARGB = Just "CCFFF" }
                                              , _fillPatternType = Just PatternTypeSolid } }
sectFill :: Maybe Fill
sectFill = Just def { _fillPattern = Just def { _fillPatternBgColor = Just def { _colorARGB = Just "FFCCC" }
                                              , _fillPatternFgColor = Just def { _colorARGB = Just "FFCCC" }
                                              , _fillPatternType = Just PatternTypeSolid } }
evenFill :: Maybe Fill
evenFill = Just def { _fillPattern = Just def { _fillPatternBgColor = Just def { _colorARGB = Just "FF9933" }
                                              , _fillPatternFgColor = Just def { _colorARGB = Just "FF9933" }
                                              , _fillPatternType = Just PatternTypeSolid } }

formatCell val fill bold = def { _formattedCell = def { _cellValue = Just $ CellText $ T.pack val }
                               , _formattedFormat = def { _formatFill = fill 
                                                        , _formatFont = Just def { _fontBold = Just bold } } }

cellFormats :: [ FormattedCell ]
cellFormats = [ def { _formattedCell = def { _cellValue = Just $ CellText "foo" }
                    , _formattedFormat = def { _formatFill = Just def { _fillPattern = Just def { _fillPatternBgColor = Just def { _colorARGB = Just "CCFFFF" }, _fillPatternFgColor = Just def { _colorARGB = Just "CCFFFF" }, _fillPatternType = Just PatternTypeSolid } } }
                   , _formattedColSpan = 1
                   , _formattedRowSpan = 1
                   }
              ]


addMenu :: [[String]] -> Map SqRef  DataValidation
addMenu rows = M.fromList $ map addMenu' $ zip [1..] rows

qregex = "^[ \t]*[A-Z][ \t]*$" :: String -- Single letter indicates a question.
-- Anything that doesn't match one of the prefixes gets a menu.
addMenu' :: (Int, [String]) -> (SqRef, DataValidation)
addMenu' (n, h:_) | (h =~ qregex) = (SqRef [singleCellRef (n, 4)], defMenu) -- these get dropdown
addMenu' (n, _) = (SqRef [singleCellRef (n, 4)], def) -- these get defaul


defMenu :: DataValidation
defMenu = def { _dvAllowBlank       = False
              , _dvError            = Just "incorrect data"
              , _dvErrorStyle       = ErrorStyleInformation
              , _dvErrorTitle       = Just "error title"
              , _dvPrompt           = Just "enter data from menu"
              , _dvPromptTitle      = Just "prompt title"
              , _dvShowDropDown     = False -- this has the opposite effect to what I expect
              , _dvShowErrorMessage = True
              , _dvShowInputMessage = True
              , _dvValidationType   = ValidationTypeList ["aaaa","bbbb","cccc"]
              }
