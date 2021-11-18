{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE PatternGuards #-} -- for string prefix pattern
module Codec.Xlsx.SimpleFormatted (makeCell, formattedCell, noBorder, lineBorder, SimpleFormat(..)) where
import Codec.Xlsx
import Codec.Xlsx.Formatted
import Control.Lens
import qualified Data.ByteString.Lazy as L
import Data.Default.Class
import Data.List (stripPrefix)
import Data.Map (Map)
import qualified Data.Map as M
import qualified Data.Text as T
import Data.Time.Clock.POSIX
import Text.CSV 
import Text.Regex.PCRE 



data SimpleFormat = SimpleFormat {
  border :: Maybe Border
  , fill :: T.Text              -- RGB color spec
  , indent :: Int
  , bold :: Bool
  , color :: T.Text
  , font :: T.Text
  }
                    
instance Default SimpleFormat where
  def = SimpleFormat { border = Nothing
                     , fill = "FFFFFF"
                     , indent = 0
                     , bold = False
                     , color = "000000"
                     , font = "Calibri"
                     }
  
class MakeCell val where
-- makeCell row col val fmt
  makeCell :: Int -> Int -> val -> SimpleFormat -> ((Int, Int), FormattedCell)
--  formatCell val fill bold border indent
  formatCell :: val -> SimpleFormat -> FormattedCell
  formatCell' :: val -> T.Text -> Bool -> Maybe Border -> Int  -> FormattedCell

instance MakeCell String where 
  makeCell rnum cnum val fmt = ((rnum, cnum), formatCell val fmt)
  formatCell val fmt = def { _formattedCell = def { _cellValue = Just $ CellText $ T.pack val }
                                               , _formattedFormat = def { _formatFill = fillColor (fill fmt)
                                                                        , _formatFont = Just def { _fontBold = Just (bold fmt)
                                                                                                 , _fontName = Just (font fmt)
                                                                                                 , _fontColor = Just $ def { _colorARGB = Just (color fmt) }
                                                                                                 }
                                                                        , _formatBorder = (border fmt)
                                                                        , _formatAlignment = Just def { _alignmentIndent = Just (indent fmt)
                                                                                                      , _alignmentHorizontal = Just CellHorizontalAlignmentLeft
                                                                                                      }
                                                                        }
                                               }
  formatCell' val fill bold border indent = def { _formattedCell = def { _cellValue = Just $ CellText $ T.pack val }
                                               , _formattedFormat = def { _formatFill = fillColor fill 
                                                                        , _formatFont = Just def { _fontBold = Just bold
                                                                                                 , _fontName = Just "Calibri"
                                                                                                 , _fontColor = Just $ def { _colorARGB = Just "000000" }
                                                                                                 }
                                                                        , _formatBorder = border
                                                                        , _formatAlignment = Just def { _alignmentIndent = Just indent
                                                                                                      , _alignmentHorizontal = Just CellHorizontalAlignmentLeft
                                                                                                      }
                                                                        }
                                               }



fillColor :: T.Text -> Maybe Fill
fillColor c = Just def { _fillPattern = Just def { _fillPatternBgColor = Just def { _colorARGB = Just c }
                                                 , _fillPatternFgColor = Just def { _colorARGB = Just c }
                                                 , _fillPatternType = Just PatternTypeSolid } }


noBorder :: Maybe Border
noBorder =  Just $ def { _borderLeft = noBorderStyle
                       , _borderRight = noBorderStyle
                       , _borderTop = noBorderStyle
                       , _borderBottom = noBorderStyle
                       }
lineBorder :: Maybe Border
lineBorder = Just $ def { _borderLeft = lineBorderStyle
                        , _borderRight = lineBorderStyle
                        , _borderTop = lineBorderStyle
                        , _borderBottom = lineBorderStyle
                        }


noBorderStyle :: Maybe BorderStyle
noBorderStyle = Just $ def { _borderStyleLine = Just LineStyleNone, _borderStyleColor = Just def { _colorARGB = Just "000000" } }

lineBorderStyle :: Maybe BorderStyle
lineBorderStyle = Just $ def { _borderStyleLine = Just LineStyleThin, _borderStyleColor = Just def { _colorARGB = Just "000000" } }


