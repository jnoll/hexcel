{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE OverloadedStrings #-}
module Data.YamlRubric (getYAML) where
-- functions for reading YAML rubrics

import qualified Data.Aeson as Aeson
import qualified Data.Aeson.KeyMap as KeyMap
import Data.Maybe (isJust)
import Data.Scientific (FPFormat(Generic), formatScientific)
import qualified Data.Text as T
import qualified Data.Vector as V hiding (concat, foldl', map, zip)
import qualified Data.Yaml as Y
import GHC.Generics
--import qualified System.Console.CmdArgs as Cmd --hiding (def)
--import System.Exit (die, exitFailure, exitSuccess, ExitCode(..))
--import Text.CSV
--import Text.Printf (printf)
--import Text.Pretty.Simple (pPrint, pPrintString)
--import Text.Regex.PCRE 



maybeValToString :: Maybe Y.Value -> String
maybeValToString Nothing = ""
maybeValToString (Just (Y.Null)) = ""
maybeValToString (Just (Y.String s)) = T.unpack s
maybeValToString (Just (Y.Number n)) = (formatScientific Generic Nothing n)
maybeValToString (Just (Y.Bool True)) = "True"
maybeValToString (Just (Y.Bool False)) = "False"
maybeValToString (Just (Y.Object o)) = undefined
maybeValToString (Just (Y.Array a)) = undefined


yamlToRow :: Y.Object -> [String]
yamlToRow els | isJust $ KeyMap.lookup "Q" els =
  let p = maybeValToString $ KeyMap.lookup "part"    els
      n = maybeValToString $ KeyMap.lookup "num"     els
      q = maybeValToString $ KeyMap.lookup "Q"       els
      a = maybeValToString $ KeyMap.lookup "A"       els
      c = maybeValToString $ KeyMap.lookup "comment" els
  in [p, n, a, c]

yamlToRow els | isJust $ KeyMap.lookup "value" els =
  let f = maybeValToString $ KeyMap.lookup "value" els
      v = maybeValToString $ KeyMap.lookup "contents" els
  in ["x", f, v]

yamlToRow _ = ["WTF?"]
  
yamlEl :: [[String]] -> (Aeson.Key, Y.Value) -> [[String]]
yamlEl result ("Part", n) = result
yamlEl result ("questions", (Y.Array a)) =
  let els = V.toList a in 
    (map yamlToRow $ map (\(Y.Object o) -> o) els) ++ result

yamlToVals :: [[String]] -> Y.Value ->  [[String]]
yamlToVals result (Y.Array a) = concat $ map (yamlToVals result) (V.toList a)
yamlToVals result (Y.Object m) = concat $ map (yamlEl result) (KeyMap.toList m)
yamlToVals _ _ = undefined

getYAML :: Y.Value -> [[String]]
getYAML y = yamlToVals [] y


