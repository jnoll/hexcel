{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE PatternGuards #-} -- for string prefix pattern
{-# LANGUAGE DeriveDataTypeable #-} -- for cmdargs
module Main where
import Codec.Xlsx
import Codec.Xlsx.SimpleFormatted 
import Codec.Xlsx.CSVRubric as CSV (makeCells, formattedCells, addMenu, makeConditionals, yesFillPattern, noFillPattern)
import Codec.Xlsx.XMLRubric as XML (procXML)
import Control.Lens
import qualified Data.ByteString.Lazy as L
import qualified Data.Map as M
import Data.List (intercalate, concat)
import Data.Time.Clock.POSIX
import System.Console.CmdArgs hiding (def)
import System.Environment (getArgs, getProgName)
import System.FilePath (takeExtension, takeBaseName, (</>), (<.>))
import System.IO (openFile, hClose, hGetContents, IOMode(..), stdin)
import Text.CSV 

-- Common to all modes

colWidth :: Int -> Double -> ColumnsProperties
colWidth col width = ColumnsProperties { cpMin = col, cpMax = col, cpWidth = Just width, cpStyle = Nothing, cpHidden = False, cpCollapsed = False, cpBestFit = False }
colWidths :: [ColumnsProperties]
colWidths = [ colWidth 1  4.0 -- section
            , colWidth 2 10.0 -- question num
            , colWidth 3 65.0 -- question
            , colWidth 4  5.0 -- answer
            , colWidth 5 12.0 -- goto if no
            , colWidth 6 15.0 -- goto if yes
            , colWidth 7 45.0 -- comment
            ]


-- Main entry point
getCSV :: String ->  [[String]]
getCSV contents = case parseCSV "/dev/stderr" contents of
    Left s -> [["error", show $ s]]
    Right rows -> rows

data Options = Options {
  opt_rubric :: FilePath
  , opt_output :: FilePath
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
  opt_rubric = "rubric.csv"    &= typFile &= help "Rubric, in CSV or XML format" &= name "rubric"
  , opt_output = "rubric.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  }
  &= summary "hexcel v0.1, (C) 2021 John Noll"
  &= program "main"


main :: IO ()
main = do
  args <- cmdArgs defaultOptions
  contents <- readFile (opt_rubric args) 

  ct <- getPOSIXTime

  case takeExtension (opt_rubric args) of
    ".csv" -> let rubric = getCSV contents
                  cells = M.fromList $ CSV.makeCells rubric
                  fmt = CSV.formattedCells cells
                  (n, ws) = head $ _xlSheets fmt
                  validations = CSV.addMenu rubric
                  conds = CSV.makeConditionals rubric
    --      formulaCells = addFormulas rubric "IF(\"Yes\", \"No\", \"Yes\")" $ _wsCells ws
                  xlsx = fmt & atSheet "Sheet1" ?~ ws { _wsDataValidations = validations
                                                      , _wsColumnsProperties = colWidths
                                                      , _wsConditionalFormattings = M.fromList conds
                                                       --, _wsCells = formulaCells
                                                      }
                  styleSheet = case parseStyleSheet $ _xlStyles xlsx of
                                 Left e -> minimalStyleSheet
                                 Right s -> s  { _styleSheetDxfs = [ CSV.yesFillPattern, CSV.noFillPattern ] } 
              in L.writeFile (opt_output args) $ fromXlsx ct xlsx { _xlStyles = renderStyleSheet styleSheet}
    ".xml" -> let result = XML.procXML contents
                  cells = M.fromList result
                  fmt = CSV.formattedCells cells
                  (n, ws) = head $ _xlSheets fmt
                  xlsx = fmt & atSheet "Sheet1" ?~ ws { _wsColumnsProperties = colWidths
                                                        --, _wsDataValidations = validations
                                                        --, _wsConditionalFormattings = M.fromList conds
                                                        --, _wsCells = formulaCells
                                                      }
                  styleSheet = case parseStyleSheet $ _xlStyles xlsx of
                                 Left e -> minimalStyleSheet
                                 Right s -> s  { _styleSheetDxfs = [ CSV.yesFillPattern, CSV.noFillPattern ] } 

               in L.writeFile (opt_output args) $ fromXlsx ct xlsx { _xlStyles = renderStyleSheet styleSheet}
    otherwise -> putStrLn $ "unknown rubric extension: " ++ (opt_rubric args)
