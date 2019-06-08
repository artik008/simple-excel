{-# LANGUAGE OverloadedStrings  #-}
{-# LANGUAGE RecordWildCards    #-}
{-# LANGUAGE StandaloneDeriving #-}
module Instances where

import qualified Data.ByteString.Lazy    as BSL
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
      simpleTag ("fgColor", [("rgb", show fillFgColor)])
    <++
      simpleTag ("bgColor", [("rgb", show fillBgColor)])
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
      , [ ("customHeight", boolToSInt sfCustomHeight)
        , ("defaultColWidth", show sfDefaultColWidth)
        , ("defaultRowHeight", show sfDefaultRowHeight)
        , ("outlineLevelCol", show sfOutlineLevelCol)
        , ("outlineLevelRow", show sfOutlineLevelRow)
        -- , ("x14ac:dyDescent", show sfX14Descent)
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

instance ToExcelXML PageSetupPr where
  toTag PageSetupPr{..} =
    simpleTag ("pageSetUpPr", [("fitToPage", show psprFitToPage)])

instance ToExcelXML PageMargins where
  toTag PageMargins{..} =
    simpleTag
      ( "pageMargins"
      , [ ("bottom", show pmBottom)
        , ("footer", show pmFooter)
        , ("header", show pmHeader)
        , ("left"  , show pmLeft)
        , ("right" , show pmRight)
        , ("top"   , show pmTop)
        ]
      )

instance ToExcelXML PageSetup where
  toTag PageSetup{..} =
    simpleTag
      ( "pageSetup"
      , [ ("firstPageNumber", show psFirstPageNumber)
        , ("fitToHeight", show psFitToHeight)
        , ("fitToWidth", show psFitToWidth)
        , ("orientation", show psOrientation)
        , ("pageOrder", show psPageOrder)
        , ("scale", show psScale)
        , ("useFirstPageNumber", boolToSInt psUseFirstPageNumber)
        ]
      )



instance ToExcelXML AppVariant where
  toTag AppVariant{..} =
    valueTag "vt:variant"
      (valueTag ("vt:" ++ ((drop 3).showToLower) variantType) variantValue)

-- | Rows and Columns

instance ToExcelXML ColumnConfig where
  toTag ColumnConfig{..} =
    simpleTag
      ( "col"
      , [ ("bestFit", showToLower columnBestFit)
        , ("customWidth", showToLower columnCustomWidth)
        , ("max", show columnLastColumn)
        , ("min", show columnFirstColumn)
        -- , ("style", show columnStyle)
        , ("width", show columnWidth)
        ]
      )

instance ToExcelXML IndRow where
  toTag IndRow{..} =
    simpleTag
      ("row",
       [ ("collapsed", showToLower rowCollapsed)
       , ("customFormat", showToLower rowCustomFmt)
       , ("customHeight", boolToSInt rowCustomHeight)
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


-- | XMLTag

instance ToExcelXML Text where
  toTag = XMLText

instance ToExcelXML Bool where
  toTag = XMLBool

instance ToExcelXML Int where
  toTag = XMLInt

instance ToExcelXML Float where
  toTag = XMLFloat

instance ToExcelXML XMLTag where
  toTag = id
