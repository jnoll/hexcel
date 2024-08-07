{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE DeriveDataTypeable #-} -- for cmdargs
module Main where
import qualified Data.Aeson as Aeson
import qualified Data.Aeson.KeyMap as KeyMap
--import qualified Data.ByteString as BS
import qualified Data.ByteString.Char8 as LU 
import Data.YamlRubric (getYAML)
import Data.List (intercalate)
--import qualified Data.HashMap.Strict as HM
import Data.Maybe (catMaybes, fromMaybe, fromJust, isJust)
--import Data.Scientific (FPFormat(Generic), formatScientific)
--import qualified Data.Text as T
--import qualified Data.Vector as V hiding (concat, foldl', map, zip)
import qualified Data.Yaml as Y -- only needed for "invert"
import GHC.Generics
import qualified System.Console.CmdArgs as Cmd --hiding (def)
import System.Exit (die, exitFailure, exitSuccess, ExitCode(..))
import Text.CSV
import Text.Printf (printf)
import Text.Pretty.Simple (pPrint, pPrintString) -- for debugging
import Text.Regex.PCRE 

  
data Row =  Part {
       pName :: String
  }
  | Question {
      qPart :: String
      , qID :: String
      , qText :: String
      , qVal :: String
      }
  | Val {
      vName :: String
      , vVal :: String
      }
  | Other {
      oText :: String
      } deriving (Generic, Show);

getCSV :: String -> ([String], [[String]])
getCSV contents = case parseCSV "contents" contents of
    Left s -> (["error"], [[show $ s]])
    Right (head:rows) -> (head, rows)

readRubric :: [[String]] -> [Row]
readRubric rs = Data.Maybe.catMaybes $ map readRow rs

-- Make a string into a regular expression.  Only to make type clear.
rx :: String -> String
rx s = s::String

readRow :: [String] -> Maybe Row
readRow ("x":n:v:_) = Just $ Val { vName = n, vVal = v }
readRow ("Section":v:_) = Nothing
readRow (p:_) | (p =~ rx "^Part ") = Just $ Part { pName = p }
readRow (p:n:q:v:_) | (p =~ rx "^[A-Z]") = Just $ Question { qPart = p, qID = n, qText = q, qVal = v }
readRow ("":v:_) = Nothing
readRow _  = Nothing 


writeCSV :: [[String]] -> [String]
writeCSV rs = map writeCSVRow rs

writeCSVRow :: [String] -> String
writeCSVRow cs = intercalate "," $ map (\c -> printf "\"%s\"" c) cs

writeYaml :: [Row] -> [String]
writeYaml rs = map writeRow rs

writeRow :: Row -> String
writeRow (Part p) = let (_:n:_) = words p in printf "- Part: %s\n  questions:" n
writeRow (Question p id txt val) =           printf "  - Q: \"%s\"\n    A: \"%s\"\n    comment: |\n    part: \"%s\"\n    num: \"%s\"" txt val p id 
writeRow (Val n v) =                         printf "  - value: %s\n    contents: |\n      %s" n $ intercalate "\n      " $ lines v  
writeRow (Other o) =                         printf "  - other: %s" o

data CmdOptions = CmdOptions {
  opt_invert :: Bool
} deriving (Cmd.Data, Cmd.Typeable, Show)

defaultCmdOptions :: CmdOptions
defaultCmdOptions = CmdOptions {
  opt_invert = False Cmd.&= Cmd.typ "Bool" Cmd.&= Cmd.help "Is input a rubric (in YAML format) to be converted to CSV?" Cmd.&= Cmd.name "invert"
  }
  Cmd.&= Cmd.summary "rupdate v0.1, (C) 2021 John Noll"
  Cmd.&= Cmd.program "main"

main :: IO ()
main = do
  args <- Cmd.cmdArgs defaultCmdOptions
  c <- getContents
  if (opt_invert args) == True then
    case (Y.decodeEither' $ LU.pack c) of
      Left err -> do die $ Y.prettyPrintParseException err
      Right y  -> do --pPrint (y :: Y.Value) -- $ yamlToVals y
        putStrLn $ printCSV $ getYAML y
    else let (h, rs) = getCSV c in do
           putStrLn "#-*-rubric-*-"
           putStrLn $ intercalate "\n" $ writeYaml $ readRubric rs
