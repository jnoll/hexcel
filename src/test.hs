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
import Text.Printf (printf)

toInt :: String -> Int
toInt s = read s

replaceVal :: String -> Cell -> Maybe Cell
replaceVal v c = Just c { _cellValue = Just $ CellText $ T.pack v }

updateCell :: Int -> Int -> String -> (CellMap, Map Int RowProperties) -> (CellMap, Map Int RowProperties)
updateCell rnum cnum v (cmap, rpmap) =
  let cells = M.update (replaceVal v) (rnum, cnum) cmap
--      rows  = M.update (setRowHeight $ fromIntegral $ length $ lines v) rnum rpmap
      rows = rpmap
  in (cells, rows)


-- experiment to see if rubric can use formulas to chnage value/appearance of cells
replaceFormula :: T.Text -> Cell -> Maybe Cell
replaceFormula f c = Just c { _cellFormula = Just $ CellFormula { _cellfAssignsToName = False, _cellfCalculate = False,  _cellfExpression = NormalFormula $ Formula { unFormula = f } } }

updateFormula :: Int -> Int -> T.Text -> CellMap -> CellMap
updateFormula rnum cnum formula cmap = M.update (replaceFormula formula) (rnum, cnum) cmap
  

createCondition :: Int -> Int -> T.Text -> Int -> (SqRef, [CfRule])
createCondition r c formula style = (SqRef [singleCellRef (r, c)], [CfRule { _cfrCondition = Expression $ Formula { unFormula = formula }, _cfrDxfId = Just (style - 1), _cfrPriority = 1, _cfrStopIfTrue = Just False } ])


setRowHeight :: Double -> RowProperties -> Maybe RowProperties
setRowHeight height rp = Just rp { rowHeight = Just $ AutomaticHeight $ height * 100.0 } 
  



-- Main entry point

data Options = Options {
  opt_values :: FilePath
  , opt_formula :: T.Text
  , opt_sheet :: T.Text
  , opt_input :: FilePath
  , opt_output :: FilePath
} deriving (Data, Typeable, Show)

defaultOptions :: Options
defaultOptions = Options {
  opt_values = "-" &= typFile &= help "Text for text area" &= name "values"
  , opt_formula = "" &= typ "String" &= help "Text for formula" &= name "formula"
  , opt_sheet = "Sheet1" &= typ "String" &= help "Sheet name to update" &= name "sheet"
  , opt_input = "empty.xlsx" &= typFile &= help "Input XLSX file name" &= name "input"
  , opt_output = "example.xlsx" &= typFile &= help "Output XLSX file name" &= name "output"
  }
  &= summary "texcel v0.1, (C) 2021 John Noll"
  &= program "main"

cond :: Int -> T.Text
cond r = T.pack $ printf "C%d=\"foo\"" r

cond' :: Int -> T.Text
cond' r = T.pack $ printf "C%d=\"bar\"" r


main :: IO ()
main = do
  args <- cmdArgs defaultOptions
  ct <- getPOSIXTime
  -- "-v -" means read values from stdin.  Useful for values provided by a pipeline.
  valHandle <- if (opt_values args) == "-" then return stdin else openFile (opt_values args) ReadMode
  contents <- hGetContents valHandle
  putStrLn contents
  excel <- L.readFile (opt_input args)

  let xlsx = toXlsx excel
      formula = (opt_formula args)
      sheet = fromJust (xlsx ^? ixSheet (opt_sheet args) )
      (cells', rowprops) = updateCell 2 3 "1" $ updateCell 3 3 contents  ((_wsCells $ sheet), (_wsRowPropertiesMap $ sheet))
--      (cells', rowprops) = updateCell 2 3 "bar" ((_wsCells $ sheet), (_wsRowPropertiesMap $ sheet))
      cells = case formula of
                "" -> cells'
                --otherwise -> updateFormula 3 4 formula $ _wsCells $ sheet
                otherwise -> updateFormula 3 4 formula $ cells'
      conds = (map (\x -> createCondition x 5 (cond x) x) $ [1..25]) ++
              (map (\x -> createCondition x 5 (cond' x) x) $ [1..25])
      names = DefinedNames [("foo", Nothing, "D12"), ("bar", Just "Sheet1", "$C$2")]
      tables = [Table {tblDisplayName = "aTable", tblName = Just "aTable", tblRef = CellRef {unCellRef = "B3"}, tblColumns = [ TableColumn {tblcName = "unused" } ], tblAutoFilter = Nothing } ]
      -- [createCondition 1 5 (cond 1) 1, createCondition 2 5 (cond 2) 2, createCondition 3 5 (cond 3) 3, createCondition 4 5 (cond 4) 4, createCondition 5 5 (cond 5) 5 ]
      xlsx' = xlsx & atSheet (opt_sheet args) ?~ sheet { _wsCells = cells
                                                       , _wsRowPropertiesMap = rowprops
                                                       , _wsConditionalFormattings = M.fromList conds
                                                       , _wsTables = tables
                                                         }
--      xlsx' = xlsx & atSheet (opt_sheet args) ?~ sheet { _wsCells = cells' }
      styleSheet = def { _styleSheetDxfs = [ def {_dxfFill = Just $ Fill { _fillPattern = Just $ def { _fillPatternBgColor = Just $ def { _colorARGB = Just "FF6666" } } } }
                                           , def {_dxfFill = Just $ Fill { _fillPattern = Just $ def { _fillPatternBgColor = Just $ def { _colorARGB = Just "00FF00" } } } } ] }
      styles = renderStyleSheet styleSheet
  -- It's considered bad practice to read from and write to the name file, so
  -- write to a temp file, then rename it.
      
  (tmpNm, tmph) <- openTempFileWithDefaultPermissions "." "output.xlsx" -- becomes "outputNNN.xlsx"
  L.hPut tmph $ fromXlsx ct xlsx' { _xlStyles = styles
                                  , _xlDefinedNames = names
                                  }
  hClose tmph

  renameFile tmpNm  (opt_output args)
