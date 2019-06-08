{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE RecordWildCards   #-}
module FileCreating where

import           Data.ByteString.Lazy (ByteString)
import           Data.List            (find, intercalate, sortOn)
import           Data.Map             (Map, size, toList)
import           Data.Text            (Text, pack)
import           Data.Time

import           Defaults
import           Models
import           Styles
import           Tag
import           Utils

sheetToXml :: (IndexedWorksheet, Int) -> (String, ByteString)
sheetToXml (IndexedWorksheet{..}, wsID) =
  ("xl/worksheets/sheet" <> show wsID <> ".xml", xml)
  where
    xml = (-++) fileHeader $
      simpleTag
        ( "worksheet"
        , [ ("mc:Ignorable","x14ac")
          , ("xmlns", urlHeader <> "spreadsheetml/2006/main")
          , ("xmlns:mc", urlHeader <> "markup-compatibility/2006")
          , ("xmlns:r", urlHeader <> "officeDocument/2006/relationships")
          ]
        )
      <<<
      filterEmptyTags
      [ emptyTag "sheetPr" <<< iwsPageSetupPrs
      , dimensionToTag iwsDimension
      , emptyTag "sheetViews" <<< iwsSheetViews
      , toTag iwsSheetFormat
      , emptyTag "cols" <<< iwsColumnsConfigs
      , emptyTag "sheetData" <<< iwsRows
      , simpleTag ("mergeCells",
          if iwsMergeCells /= []
          then [("count", show $ length iwsMergeCells)] else [])
        <<< iwsMergeCells
      , toTag iwsPageMargins
      , toTag iwsPageSetup
      ]

createWorkBook :: Xlsx -> ByteString
createWorkBook Xlsx{..} = (-++) fileHeader $
    simpleTag
      ( "workbook"
      , [ ("xmlns:r", urlHeader <> "officeDocument/2006/relationships")
        , ("xmlns", urlHeader <> "spreadsheetml/2006/main")
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
    , [("xmlns", urlHeader <> "package/2006/relationships")]
    )
    <<<
    [ relSharedStrings
    , relStyles
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
    typeUrl = urlHeader <> "officeDocument/2006/relationships/worksheet"

createRels :: ByteString
createRels = (-++) fileHeader $
  simpleTag
    ( "Relationships"
    , [("xmlns", urlHeader <> "package/2006/relationships")]
    )
    <<< [relWorkBook, relCore, relApp]


createContentTypes :: Int -> ByteString
createContentTypes l = (-++) fileHeader $
  simpleTag
    ("Types"
    , [("xmlns", urlHeader <> "package/2006/content-types")]
    )
    <<< overrides
    <<< map
      (\x -> Override cType ("/xl/worksheets/sheet" <> show x <> ".xml"))
      [1..l]
  where
    cType = appPathHead <> "officedocument.spreadsheetml.worksheet+xml"

createSharedStrings :: StringHistory -> ByteString
createSharedStrings hist = (-++) fileHeader $
  simpleTag
    ("sst",[("xmlns", urlHeader <> "spreadsheetml/2006/main")])
  <<< (map stringToTag sortedHist)
  where
    sortedHist = map fst $ sortOn snd $ toList hist
    stringToTag s = emptyTag "si" <++ (valueTag "t" s)

createStylesFile :: StyleHistory -> ByteString
createStylesFile hist = (-++) fileHeader $
  simpleTag
    ("styleSheet"
    , [
        ("xmlns", urlHeader <> "spreadsheetml/2006/main")
      , ("mc:Ignorable","x14ac")
      , ("xmlns:mc", urlHeader <> "markup-compatibility/2006")
      ]
    )
  <++
    (simpleTag
      ("numFmts", countAttr ssNumFmts)
      <<< (sortMapToList ssNumFmts)
    )
  <++
    (simpleTag
      ("fonts", countAttr ssFonts)
      <<< (sortMapToList ssFonts)
    )
  <++
    (simpleTag
      ("fills", countAttr ssFills)
      <<< (sortMapToList ssFills)
    )
  <++
    (simpleTag
      ("borders", countAttr ssBorders)
      <<< (sortMapToList ssBorders)
    )
  <++
    (simpleTag
      ("cellXfs", countAttr ssCellXfs)
      <<< (map snd $ sortOn fst $ toList ssCellXfs)
    )
  -- <++ (simpleTag ("dxfs",[("count","0")]))
  where
    sortMapToList f = map fst $ sortOn snd $ toList f
    countAttr a = [("count", show $ size a)]
    StyleSheet{..} = styleHistoryToStyleSheet hist


createApp :: Xlsx -> ByteString
createApp Xlsx{..} = (-++) fileHeader $
  simpleTag
    ( "Properties"
    , [
        ("xmlns", urlHeader <> "officeDocument/2006/extended-properties")
      , ("xmlns:vt", urlHeader <> "officeDocument/2006/docPropsVTypes")
      ]
    )
  <++ valueTag "TotalTime" (0 :: Int)
  <++ valueTag "Application" ("Microsoft Excel" :: Text)
  <++ valueTag "DocSecurity" (0 :: Int)
  <++ valueTag "ScaleCrop" False
  <++ (simpleTag ("HeadingPairs", [])
        <++ simpleTag
              (
                "vt:vector"
              , [ ("baseType", "variant")
                , ("size", show $ length headingPairs)
                ]
              )
              <<< headingPairs
      )
  <++ (simpleTag ("TitlesOfParts", [])
        <++ simpleTag
              (
                "vt:vector"
              , [ ("baseType", "lpstr")
                , ("size", show $ length titlesOfParts)
                ]
              )
              <<< titlesOfParts
      )
  <++ valueTag "LinksUpToDate" False
  <++ valueTag "SharedDoc" False
  <++ valueTag "HyperlinksChanged" False
  <++ valueTag "AppVersion" (15.0300 :: Float)
  where
    wsInfo = map (pack.wsName) worksheets
    headingPairs = appHeadingPairs $ length wsInfo
    titlesOfParts = appTitlesOfParts wsInfo

createCore :: UTCTime -> ByteString
createCore utc = (-++) fileHeader $
  simpleTag
    ( "cp:coreProperties"
    , [ ("xmlns:cp", urlHeader <> "package/2006/metadata/core-properties")
      , ("xmlns:dc", "http://purl.org/dc/elements/1.1/")
      , ("xmlns:dcterms", "http://purl.org/dc/terms/")
      , ("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
      ]
    )
  <++ simpleTag ("dcterms:created", [("xsi:type", "dcterms:W3CDTF")])
        <++ (XMLText $ fmt)
  <++ valueTag "dc:creator" ("simple-excel" :: Text)
  <++ valueTag "cp:lastModifiedBy" ("simple-excel" :: Text)
  where
    fmt = pack $ formatTime
            defaultTimeLocale
            "%Y-%m-%dT%H:%M:%SZ"
            (utcToLocalTime (hoursToTimeZone 3) utc)

filterEmptyTags :: [XMLTag] -> [XMLTag]
filterEmptyTags tags = filter (\(XMLTag (_,l) b _) -> l /= [] || b /= []) tags
