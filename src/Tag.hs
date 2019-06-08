{-# LANGUAGE OverloadedStrings #-}
module Tag where

import qualified Data.ByteString.Lazy    as BSL
import           Data.List               (intercalate)
import           Data.String.Conversions (cs)


import           Models
import           Utils

(-++) :: BSL.ByteString -> XMLTag -> BSL.ByteString
(-++) bsl tag = bsl <> toXML tag

toXML :: XMLTag -> BSL.ByteString
toXML (XMLTag (name, args) body end) =
  (
    cs $
      "<" <> name <>
      concat (map (\(x,y) -> " " <> x <> "=\"" <> y <> "\"") args)
  ) <>
  if end
    then ">" <>
      BSL.intercalate "\n  " (map toXML $ reverse body) <>
      "\n</" <> cs name <> ">"
    else "/>"
toXML (XMLFloat v) = cs $ show v
toXML (XMLInt v)   = cs $ show v
toXML (XMLText v)  = cs v
toXML (XMLBool v)  = cs $ showToLower v
