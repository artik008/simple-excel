{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
module Instances where

import           Data.Char               (toLower)
import           Data.List               (find, intercalate, sortOn)
import           Data.Maybe              (catMaybes, fromJust, isJust)
import           Data.String             (IsString, fromString)
import           Data.String.Conversions (cs)
import           Data.Text               (Text)

import           Models
import           Utils

-- ** ToExcelXML

-- | Styles

instance ToExcelXML IndCellXf where
  toTag IndCellXf{..} =
    simpleTag
      ( "xf"
      , ( (checkID ("applyAlignment", ialignment):
        map checkID
        [ ("applyBorder", iborder)
        , ("applyFill", ifill)
        , ("applyFont", ifont)
        , ("applyNumberFormat", inumFmt)
        ])) ++
        catMaybes (map addName
        [ ("borderId", iborder)
        , ("fillId", ifill)
        , ("fontId", ifont)
        , ("numFmtId", inumFmt)
        ]) ++
        [ ("xfId", show icellXfIndex)]
      )
    <<< catMaybes [ialignment]
    where
      addName (n, x) = (\y -> (n, show y)) <$> x
      checkID (n, x) = (n, boolToSInt $ isJust x)

instance ToExcelXML Border where
  toTag Border{..} =
    emptyTag "border"
    <<<
    catMaybes borderStyles

instance ToExcelXML BorderStyle where
  toTag BorderStyle{..} =
    simpleTag (show borderSide, [("style", map toLower $ show borderLine)])
      <++ simpleTag ("color", [("rgb", show borderColor)])

instance IsString Color where
  fromString a = RGB $ cs a

instance ToExcelXML Fill where
  toTag Fill{..} =
    emptyTag "fill"
    <++
    ( simpleTag ("patternFill", [("patternType", cs fillType)])
    <++
      simpleTag ("fgColor", [("rgb", show fillColor)])
    )

instance ToExcelXML Font where
  toTag Font{..} =
    emptyTag "font"
    <<<
    [ simpleTag ("name", [("val", cs fontName)])
    , simpleTag ("b", [("val", boolToSInt fontBold)])
    , simpleTag ("outline", [("val", boolToSInt fontOutline)])
    , simpleTag ("sz", [("val", show fontSize)])
    , simpleTag ("color", [("rgb", show fontColor)])
    , simpleTag ("vertAlign", [("val", show fontVertAlignment)])
    ]

instance ToExcelXML NumFmt where
  toTag NumFmt{..} =
    simpleTag
      ( "numFmt"
      , [ ("formatCode", formatCode)
        , ("numFmtId", show numFmtId)
        ]
      )

instance ToExcelXML Alignment where
  toTag Alignment{..} =
    simpleTag
      ( "alignment"
      , [ ("horizontal", show horizontal)
        , ("vertical", show vertical)
        , ("wrapText", cs wrapText)
        ]
      )

-- | Additional files

instance ToExcelXML Relationship where
  toTag Relationship{..} =
    simpleTag
      ( "Relationship"
      , [ ("Id", "rId" <> show relID)
        , ("Target", relTarget)
        , ("Type", relType)
        ]
      )


instance ToExcelXML Override where
  toTag Override{..} =
    simpleTag
      ( "Override"
      , [ ("ContentType", orContentType)
        , ("PartName", orPartName)
        ]
      )

instance ToExcelXML SheetFormat where
  toTag SheetFormat{..} =
    simpleTag
      ( "sheetFormatPr"
      , [ ("customHeight", show sfCustomHeight)
        , ("defaultColWidth", show sfDefaultColWidth)
        , ("defaultRowHeight", show sfDefaultRowHeight)
        , ("outlineLevelCol", show sfOutlineLevelCol)
        , ("outlineLevelRow", show sfOutlineLevelRow)
        , ("x14ac:dyDescent", show sfX14Descent)
        ]
      )

instance ToExcelXML SheetView where
  toTag SheetView{..} =
    simpleTag
      ( "sheetView"
      , [ ("defaultGridColor", show svDefaultGridColor)
        , ("showGridLines", show svShowGridLines)
        , ("workbookViewId", show svWorkbookViewId)
        ]
      )

instance ToExcelXML PageSetup where
  toTag PageSetup{..} =
    simpleTag ("pageSetUpPr", [("fitToPage", show psFitToPage)])


-- | Rows and Columns

instance ToExcelXML ColumnConfig where
  toTag ColumnConfig{..} =
    simpleTag
      ( "col"
      , [ ("bestFit", show columnBestFit)
        , ("customWidth", show columnCustomWidth)
        , ("max", show columnLastColumn)
        , ("min", show columnFirstColumn)
        , ("width", show columnWidth)
        ]
      )

instance ToExcelXML IndRow where
  toTag IndRow{..} =
    simpleTag
      ("row",
       [ ("collapsed", showToLower rowCollapsed)
       , ("customFormat", showToLower rowCustomFmt)
       , ("customHeight", showToLower rowCustomHeight)
       , ("hidden", showToLower rowHidden)
       , ("ht", show rowHeight)
       , ("outlineLevel", show rowOutlineLevel)
       , ("r", show irowNumber)
       ])
    <<< irowCells
    where
      RowConfig{..} = irowConfig

-- | Cells

instance ToExcelXML IndCell where
  toTag IndCell{..} =
    simpleTag
      ( "c"
      , [
          ("r", coordsToName icellCoords)
        , ("s", show icellStyle)
        ] ++ cellType icellValue
      )
      <<< case icellValue of
            Just v  -> toTagList v
            Nothing -> []

instance ToExcelXML IndCellValue where
  toTagList (IndCellText t) = [valueTag "v" t]
  toTagList (IndCellFloat t) = [valueTag "v" t]
  toTagList (IndCellFormula f mv) =
    [valueTag "f" f]
    ++ if mv /= Nothing then [valueTag "v" (fromJust mv)] else []


-- | Merges

instance ToExcelXML Merge where
  toTag (Merge m) = simpleTag ("mergeCell", [("ref", coordsToSRange m)])

instance ToExcelXML Text where
  toTag = XMLText

instance ToExcelXML Int where
  toTag = XMLInt

instance ToExcelXML Float where
  toTag = XMLFloat

instance ToExcelXML XMLTag where
  toTag = id
