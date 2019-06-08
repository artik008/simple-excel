{-# LANGUAGE RecordWildCards #-}
module Indexing where

import           Data.Map   hiding (filter, foldl, map)
import qualified Data.Map   as Map
import           Data.Maybe (fromJust, fromMaybe)
import           Data.Tuple
import           Prelude    hiding (lookup)


import           Defaults
import           Models

addIndexedCell
  :: WithHistory IndCell
  -> (CellCoords, Cell)
  -> WithHistory IndCell
addIndexedCell (WithHistory l styhist strhist) (cellCoords, Cell{..}) =
  WithHistory
    (ic:l)
    (insert cellStyle newStyID styhist)
    (fromMaybe strhist $ (\x ->
      case x of
        CellText t -> (insert t newStrID strhist)
        _          -> strhist
      ) <$> cellValue
    )
  where
    newStyID =
      fromMaybe
        ((+) 1 $ maximum $ (-1:elems styhist))
        (lookup cellStyle styhist)
    newStrID =
      case cellValue of
        Just v -> case v of
          CellText t ->
            fromMaybe
              ((+) 1 $ maximum $ (-1:elems strhist))
              (lookup t strhist)
          _ -> 0
        Nothing -> 0
    ic = IndCell
      { icellValue  = (\x -> case x of
          CellText t      -> IndCellText newStrID
          CellFloat v     -> IndCellFloat v
          CellFormula f v -> IndCellFormula f v) <$> cellValue
      , icellStyle  = newStyID
      , icellCoords = cellCoords
      }

addIndexedRow
  :: WithHistory IndRow
  -> (Int, RowConfig, [(CellCoords, Cell)])
  -> WithHistory IndRow
addIndexedRow (WithHistory l styhist strhist) (rnum, conf, cells) =
  WithHistory
    (ir:l)
    newStyHist
    newStrHist
  where
    ir = IndRow
      { irowConfig = conf
      , irowNumber = rnum
      , irowCells  = reverse indCells
      }
    WithHistory indCells newStyHist newStrHist =
      foldl addIndexedCell (WithHistory [] styhist strhist) cells

addIndWorkSheet
  :: WithHistory IndexedWorksheet
  -> Worksheet
  -> WithHistory IndexedWorksheet
addIndWorkSheet (WithHistory l indexedStyles sharedStrings) Worksheet{..} =
  WithHistory (iws:l) newIndStyles newSharedStrings
  where
    iws = IndexedWorksheet
      { iwsName           = wsName
      , iwsDimension      = dim
      , iwsRows           = reverse indexedRows
      , iwsColumnsConfigs = wsColumnsConfigs
      , iwsMergeCells     = fMergeCells
      , iwsPageSetupPrs   = wsPageSetupPrs
      , iwsSheetViews     = wsSheetViews
      , iwsSheetFormat    = wsSheetFormat
      , iwsPageMargins    = wsPageMargins
      , iwsPageSetup       = wsPageSetup
      }
    WithHistory indexedRows newIndStyles newSharedStrings =
      foldl addIndexedRow (WithHistory [] indexedStyles sharedStrings) rowsList
    dim@((minC, minR), (maxC, maxR)) = countDimension activeCells wsStyleFix
    -- rdim@((minC, minR), (maxC, maxR)) = countDimension reversedActiveCells wsStyleFix
    activeDim = countDimension activeCells False
    -- ractiveDim = countDimension reversedActiveCells False
    activeCells = ((keys wsCells)) ++
      (concat $ map (\(Merge (c1,c2)) -> [c1, c2]) fMergeCells)
    -- reversedActiveCells = map swap activeCells
    -- reversedMergeCells = map (\(Merge (c1, c2)) -> Merge (swap c1, swap c2)) fMergeCells
    rowsList =
      map
        (\(num, conf) ->
          ( num, conf
          , map (\col -> ((col, num),filledCells!(col, num))) [minC..maxC]
          )
        ) $
        toList $
          foldl (\x y -> insert y defRowConfig x) wsRowsConfig [minR..maxR] -- FIXME
    filledCells = fillCells wsCells fMergeCells dim activeDim
    -- reversedCells = Map.foldlWithKey
      -- (\cells (x,y) c -> Map.insert (y,x) c cells) Map.empty filledCells
    fMergeCells = filter (\(Merge (f,l)) -> f /= l) wsMergeCells

fillCells
  :: Map CellCoords Cell
  -> [Merge]
  -> Range
  -> Range
  -> Map CellCoords Cell
fillCells cells merges range actRange =
    foldl setMergeStyle
    (foldl
      (\oldc1 y -> foldl
        (\oldc2 x ->
          if notMember (x, y) oldc2
          then insert (x, y) (Cell Nothing (chooseStyle (x, y))) oldc2
          else oldc2
        )
        oldc1
        [minC..maxC]
      )
      cells
      [minR..maxR]
    )
    merges
  where
    ((minC, minR), (maxC, maxR)) = range
    ((aminC, aminR), (amaxC, amaxR)) = actRange
    chooseStyle (x, y)
      | x >= aminC && x <= amaxC && y == (aminR - 1) = defTopStyle
      | x >= aminC && x <= amaxC && y == (amaxR + 1) = defBottomStyle
      | y >= aminR && y <= amaxR && x == (aminC - 1) = defLeftStyle
      | y >= aminR && y <= amaxR && x == (amaxC + 1) = defRightStyle
      | x <= amaxC && x >= aminC && y <= amaxR && y >= aminR = defStyle
      | otherwise = defOutStyle

setMergeStyle :: Map CellCoords Cell -> Merge -> Map CellCoords Cell
setMergeStyle cells (Merge ((mC1, mR1), (mC2, mR2))) =
  foldl
    (\oldc1 y -> foldl
      (\oldc2 x -> changeStyle x y oldc2)
      oldc1
      [minC..(max mC1 mC2)]
    )
    cells
    [minR..(max mR1 mR2)]
  where
    mergeStyle = cellStyle $ fromJust $ lookup (minC, minR) cells
    minR = min mR1 mR2
    minC = min mC1 mC2
    changeStyle x y oldcells = if (x,y) == (minC, minR) then oldcells else
      let c = oldcells!(x,y)
        in insert (x,y) c{cellStyle = mergeStyle, cellValue = Nothing} oldcells

countDimension :: [CellCoords] -> Bool -> Range
countDimension [] _ = ((1,1),(5,10)) -- FIXME
countDimension cells fix =
  ( ( checkRange 1 minimum fst, checkRange 1 minimum snd)
  , (checkRange 5 maximum fst, checkRange 10 maximum snd)
  )
  where
    checkRange def minmax cell =
      let first = if fix then [def] else []
        in minmax $ first ++ (map cell cells)
