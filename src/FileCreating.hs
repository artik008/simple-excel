{-# LANGUAGE RecordWildCards #-}
module FileCreating where

import           Data.ByteString.Lazy (ByteString)
import           Data.List            (find, intercalate, sortOn)
import           Data.Map             (Map, size, toList)

import           Defaults
import           Models
import           Styles
import           Utils

sheetToXml :: (IndexedWorksheet, Int) -> (String, ByteString)
sheetToXml (IndexedWorksheet{..}, wsID) =
  ("xl/worksheets/sheet" <> show wsID <> ".xml", xml)
  where
    xml = (-++) fileHeader $
      simpleTag
        ( "worksheet"
        , [ ("xmlns","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
          , ("xmlns:r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
          ]
        )
      <<<
      [ emptyTag "sheetPr" <<< iwsPageSetups
      , dimensionToTag iwsDimension
      , emptyTag "sheetViews" <<< iwsSheetViews
      , toTag iwsSheetFormat
      , emptyTag "cols" <<< iwsColumnsConfigs
      , emptyTag "sheetData" <<< iwsRows
      , emptyTag "mergeCells" <<< iwsMergeCells
      ]

createWorkBook :: Xlsx -> ByteString
createWorkBook Xlsx{..} = (-++) fileHeader $
    simpleTag
      ( "workbook"
      , [ ("xmlns:r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        , ("xmlns","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        ]
      )
    <<<
    [ emptyTag "workbookPr"
    , emptyTag "bookViews" <++ emptyTag "workbookView"
    , emptyTag "sheets" <<< (map worksheetToWBTag $ zip [1..] worksheets)
    , emptyTag "definedNames"
    ]

worksheetToWBTag :: (Int, Worksheet) -> XMLTag
worksheetToWBTag (wsID, Worksheet{..}) =
  simpleTag
    ( "sheet"
    , [ ("name", wsName)
      , ("sheetId", show wsID)
      , ("r:id", "rId" <> show (1000 + wsID))
      ]
    )

createWSRels :: Int -> ByteString
createWSRels wsNum = (-++) fileHeader $
    simpleTag
      ( "Relationships"
      , [("xmlns","http://schemas.openxmlformats.org/package/2006/relationships")]
      )
      <<<
      [ relSharedStrings
      , relStyles
      , relTheme
      ]
      <<<
      map
        (\x -> Relationship
          { relID = x + 1000
          , relTarget = "worksheets/sheet" <> show x <> ".xml"
          , relType = typeUrl
          }
        ) [1..(wsNum)]
  where
    typeUrl = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

createContentTypes :: Int -> ByteString
createContentTypes l = (-++) fileHeader $
  simpleTag
    ("Types"
    , [("xmlns","http://schemas.openxmlformats.org/package/2006/content-types")]
    )
    <<< overrides
    <<< map
      (\x -> Override cType ("/xl/worksheets/sheet" <> show x <> ".xml"))
      [1..l]
  where
    cType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

createSharedStrings :: StringHistory -> ByteString
createSharedStrings hist = (-++) fileHeader $
  simpleTag
    ("sst",[("xmlns","http://schemas.openxmlformats.org/spreadsheetml/2006/main")])
  <<< (map stringToTag sortedHist)
  where
    sortedHist = map fst $ sortOn snd $ toList hist
    stringToTag s = emptyTag "si" <++ (valueTag "t" s)

createStylesFile :: StyleHistory -> ByteString
createStylesFile hist = (-++) fileHeader $
  simpleTag
    ("styleSheet",[("xmlns","http://schemas.openxmlformats.org/spreadsheetml/2006/main")])
  <++
    (simpleTag
      ("borders", countAttr ssBorders)
      <<< (sortMapToList ssBorders)
    )
  <++
    (simpleTag
      ("fills", countAttr ssFills)
      <<< (sortMapToList ssFills)
    )
  <++
    (simpleTag
      ("fonts", countAttr ssFonts)
      <<< (sortMapToList ssFonts)
    )
  <++
    (simpleTag
      ("numFmts", countAttr ssNumFmts)
      <<< (sortMapToList ssNumFmts)
    )
  <++
    (simpleTag
      ("cellXfs", countAttr ssCellXfs)
      <<< (map snd $ sortOn fst $ toList ssCellXfs)
    )
  <++ (simpleTag ("dxfs",[("count","0")]))
  where
    sortMapToList f = map fst $ sortOn snd $ toList f
    countAttr a = [("count", show $ size a)]
    StyleSheet{..} = styleHistoryToStyleSheet hist
