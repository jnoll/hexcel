{-# LANGUAGE FlexibleInstances #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE PatternGuards #-} -- for string prefix pattern
module Codec.Xlsx.XMLRubric (procXML) where
import Codec.Xlsx
import Codec.Xlsx.Formatted (formatWorkbook, FormattedCell(..))
import Codec.Xlsx.SimpleFormatted 
import Control.Monad.State
import Data.Default
import Data.Maybe (fromJust, fromMaybe)
import qualified Data.Text as T
import Text.Printf
import Text.XML.Light
import Text.XML.Light.Proc (onlyElems)

-- Accumulator of state and emerging spreadsheet.
data Accum= Accum {
  r_line :: Int
} deriving (Eq, Show)

instance Default Accum where
  def = Accum { r_line = 1 }

-- XXXjn placeholder.  Should count sections/depth or something.
qColor :: Int -> T.Text
qColor q  = "FFCC99"

-- Helper functions to get stuff from elements
nameIs :: Element -> String -> Bool
nameIs e n = (qName $ elName e) == n


findChild' :: Element -> String -> Element
findChild' e n = fromMaybe blank_element $ findChild (blank_name { qName = n } ) e

findChildren' :: Element -> String -> [Element]
findChildren' e n = elChildren $ findChild' e n

-- Element processing
procRule :: Element ->  State Accum [((Int, Int), FormattedCell)]
procRule e = do
  a <- get
  let l = r_line a
      q = strContent $ findChild' e "q"
      f = strContent $ findChild' e "fact"
      bgColor = qColor l
      sec = "A" :: String
      num = "1.2.3" :: String
  put $ a { r_line = l + 1} 
  es <- procElems (findChildren' e "sub-rules") []
  return $ ( [ makeCell l 1 sec def { fill = bgColor }
             , makeCell l 2 num def { fill = bgColor }
             , makeCell l 3 q def { fill = bgColor }
             , makeCell l 4 ("?" :: String) def { fill = "FFFF66", border = lineBorder}
             , makeCell l 5 ("" :: String) def { fill = bgColor }
             , makeCell l 6 ("" :: String) def { fill = bgColor }
             , makeCell l 7 ("" :: String) def { fill = bgColor }
             ] ++ es)

--  return $ (printf "%d:Rule: %s (%s)" l q f:es)

procPart ::  Element -> State Accum [((Int, Int), FormattedCell)]
procPart e = do
  a <- get
  let l = r_line a
      nm = strContent $ findChild' e "name"
      bgColor = "99CCFF"
      sec = "A" :: String
  put $ a { r_line = l + 1} 
  es <- procElems (findChildren' e "rules") []
  return $ ( [ makeCell l 1 ("Part" :: String)  def { fill = bgColor }
             , makeCell l 2 sec def { fill = bgColor }
             , makeCell l 3 nm def { fill = bgColor }
             , makeCell l 4 ("" :: String) def { fill = bgColor }
             , makeCell l 5 ("" :: String) def { fill = bgColor }
             , makeCell l 6 ("" :: String) def { fill = bgColor }
             , makeCell l 7 ("" :: String) def { fill = bgColor }
             ] ++ es)
--  return $ (printf "%d:Part: %s" l nm:es)

procVal :: Element -> State Accum [((Int, Int), FormattedCell)]
procVal e = do
  a <- get
  let l = r_line a
      nm = strContent $ findChild' e "name"
      bgColor = "FFDDFF"
  put $ a { r_line = l + 1} 
  return $ [ makeCell l 1 ("x" :: String) def { fill = bgColor }
           , makeCell l 2 nm def { fill = bgColor }
           , makeCell l 3 ("" :: String) def { fill = bgColor }
           , makeCell l 4 ("" :: String) def { fill = bgColor }
           , makeCell l 5 ("" :: String) def { fill = bgColor }
           , makeCell l 6 ("" :: String) def { fill = bgColor }
           , makeCell l 7 ("" :: String) def { fill = bgColor }
           ]
--  return $ [(printf "%d:Val: %s" l nm)] -- vals are leaves, so don't try to recurse

procElem :: Element ->  State Accum [((Int, Int), FormattedCell)]
procElem e | nameIs e "part" = procPart e
procElem e | nameIs e "rule" = procRule e
procElem e | nameIs e "val" = procVal e
procElem e = procElems (onlyElems $ elContent e) [] -- >>= (\es -> return ((qName $ elName e) ++ es))

procElems :: [Element] -> [((Int, Int), FormattedCell)] -> State Accum [((Int, Int), FormattedCell)]
procElems [] ss = return ss 
procElems (e:ns) ss = procElem e >>= (\r -> procElems ns (ss ++ r))

    
procXML :: String -> [((Int, Int), FormattedCell)]
procXML c =
  let (result, state) = runState (procElems (onlyElems $ parseXML c) []) def
  in result


