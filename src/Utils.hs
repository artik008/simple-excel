module Utils where

import qualified Data.ByteString.Lazy    as BSL
import           Data.Char               (chr, ord, toLower)
import           Data.String.Conversions (cs)
import qualified Data.Text               as T

import           Models

boolToSInt :: Bool -> String
boolToSInt b = if b then "1" else "0"

dimensionToTag :: Range -> XMLTag
dimensionToTag r = simpleTag ("dimension", [("ref", coordsToSRange r)])

showToLower :: (Show a) => a -> String
showToLower = (map toLower).show

isText :: IndCellValue -> Bool
isText (IndCellText _) = True
isText _               = False

cellType :: Maybe IndCellValue -> [Attr]
cellType Nothing  = []
cellType (Just v) = [("t", if isText v then "s" else "n")]

valueTag :: (ToExcelXML a) => String -> a -> XMLTag
valueTag name t = XMLTag (name,[]) [toTag t] True

simpleTag :: TagBegin -> XMLTag
simpleTag = \x -> XMLTag x [] False

emptyTag :: String -> XMLTag
emptyTag name = XMLTag (name,[]) [] False

getName :: BSL.ByteString -> BSL.ByteString
getName xml = cs $ takeWhile (/=' ') $ drop 1 $ cs xml

coordsToName :: CellCoords -> String
coordsToName (c, r) =  (int2col c) <> (cs $ show r)

-- | convert column number (starting from 1) to its textual form (e.g. 3 -> \"C\")
int2col :: Int -> String
int2col = reverse . map int2let . base26
    where
        int2let 0 = 'Z'
        int2let x = chr $ (x - 1) + ord 'A'
        base26  0 = []
        base26  i = let i' = (i `mod` 26)
                        i'' = if i' == 0 then 26 else i'
                    in seq i' (i' : base26 ((i - i'') `div` 26))

-- | reverse to 'int2col'
col2int :: T.Text -> Int
col2int = T.foldl' (\i c -> i * 26 + let2int c) 0
    where
        let2int c = 1 + ord c - ord 'A'

coordsToSRange :: (CellCoords, CellCoords) -> String
coordsToSRange (c1, c2) = cs $ coordsToName c1 <> ":" <> coordsToName c2
