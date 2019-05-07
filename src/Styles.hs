{-# LANGUAGE RecordWildCards #-}
module Styles where

import           Data.Map

import           Models

styleHistoryToStyleSheet :: StyleHistory -> StyleSheet
styleHistoryToStyleSheet hist =
  Prelude.foldl addStyleToSheet emptyStyleSheet (toList hist)
  where
    emptyStyleSheet = StyleSheet empty empty empty empty empty

addStyleToSheet :: StyleSheet -> (CellXf, Int) -> StyleSheet
addStyleToSheet ss@StyleSheet{..} (c@CellXf{..}, num) =
  if member num ssCellXfs
  then ss
  else StyleSheet
    { ssBorders    = newBorders
    , ssFills      = newFills
    , ssFonts      = newFonts
    , ssNumFmts    = newNumFmts
    , ssCellXfs    = insert num newCellXf ssCellXfs
    }
  where
    newCellXf = IndCellXf
      { icellXfIndex = num
      , iborder      = borderId
      , ifill        = fillId
      , ifont        = fontId
      , inumFmt      = numFmtId
      , ialignment   = alignment
      }
    (newBorders, borderId) = checkStyle ssBorders border
    (newFills, fillId)     = checkStyle ssFills fill
    (newFonts, fontId)     = checkStyle ssFonts font
    (newNumFmts, numFmtId) = checkStyle ssNumFmts numFmt

checkStyle
  :: (Ord k)
  => Map k StyleID
  -> Maybe k
  -> (Map k StyleID, Maybe StyleID)
checkStyle m Nothing  = (m, Nothing)
checkStyle m (Just s) = case m !? s of
  Just sId -> (m, Just sId)
  Nothing  -> (insert s newId m, Just newId)
  where
    newId = (+) 1 $ maximum $ (-1:elems m)
