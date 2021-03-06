{-# LANGUAGE InstanceSigs      #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE QuasiQuotes       #-}
{-# LANGUAGE RecordWildCards   #-}
module Xlsx where

import qualified Codec.Archive.Zip       as Zip
import qualified Data.ByteString.Lazy    as BSL
import           Data.Char
import           Data.List               (find, intercalate, sortOn)
import           Data.Map                hiding (drop, foldl, foldr, map)
import qualified Data.Map                as Map
import           Data.Maybe
import           Data.String.Conversions
import           Data.Text               (Text)
import qualified Data.Text               as T
import           Data.Time
import           NeatInterpolation
import           System.Directory

import           Defaults
import           FileCreating
import           Indexing
import           Instances
import           Models

createXlsx :: FilePath -> Xlsx -> IO BSL.ByteString
createXlsx fname xlsx@Xlsx{..} = do
  now <- getCurrentTime
  files <- filesInDir fname
  let WithHistory indWorksheets indStyles sharedStrings =
        foldl
          addIndWorkSheet
          (WithHistory [] (singleton defStyle 0) empty)
          worksheets
      wshs = map sheetToXml $ zip (reverse indWorksheets) [1..]
  openedFiles <- sequence $
    map
      (\x -> do
        f <- BSL.readFile x
        return $ (drop (length fname + 1) x, f)
      )
      files
  let archive = foldr
        Zip.addEntryToArchive
        Zip.emptyArchive
        (map (\(x, y) -> Zip.toEntry x 5 y)
          (
            openedFiles <>
            wshs <>
            [("xl/workbook.xml", createWorkBook xlsx)
            ,("xl/_rels/workbook.xml.rels", createWSRels (length worksheets))
            ,("[Content_Types].xml", createContentTypes (length worksheets))
            ,("xl/sharedStrings.xml", createSharedStrings sharedStrings)
            ,("xl/styles.xml", createStylesFile indStyles)
            ,("docProps/app.xml", createApp xlsx)
            ,("docProps/core.xml", createCore now)
            ,("_rels/.rels", createRels)
            ]
          )
        )
  return $ Zip.fromArchive $ archive
  where
    filesInDir name = do
      inDir <- listDirectory name
      checked <- sequence $
        map (\f -> do
          isDir <- doesDirectoryExist (name <> "/" <> f)
          if isDir
          then filesInDir (name <> "/" <> f)
          else return [name <> "/" <> f]
        ) inDir
      return $ concat checked


-- | Testing


test :: IO ()
test = do
  xlsx <- createXlsx "template" testTable
  BSL.writeFile "test-report.xlsx" xlsx

headerStyle :: CellXf
headerStyle = CellXf
  { border    = Just grayBorder
  , fill      = Just $ Fill "solid" "000000" "000000"
  , font      = Just $ Font "Helvetica Neue" True False 14 (RGB "f5e14c") Baseline
  , numFmt    = Nothing
  , alignment = Just $ Alignment HCenter VCenter "1"
  }

grayBorder :: Border
grayBorder = Border
  { borderStyles =
     [
       Just $ BorderStyle BLeft     Thin (RGB "808080")
     , Just $ BorderStyle BRight    Thin (RGB "808080")
     , Just $ BorderStyle BTop      Thin (RGB "808080")
     , Just $ BorderStyle BBottom   Thin (RGB "808080")
     , Just $ BorderStyle BDiagonal Thin (RGB "808080")
     ]
  }

testTable :: Xlsx
testTable = Xlsx
  { worksheets = [ws1 "Лист 1", ws1 "Лист 2"]
  }
  where
    ws1 n = Worksheet
      { wsName           = n
      , wsStyleFix       = True
      , wsCells          = fromList
        [
          ((1,1), Cell (Just $ CellText "0") headerStyle)
        , ((2,2), Cell (Just $ CellFloat 40.2) defStyle{border = Just grayBorder})
        , ((3,3), Cell (Just $ CellFormula "B2/B2" (Just "1")) headerStyle)
        ]
      , wsRowsConfig     = fromList
          [ (1, RowConfig False True False False 30 0)
          , (2, RowConfig False True False False 50 0)
          , (3, RowConfig False True False False 30 0)
          ]
      , wsColumnsConfigs =
        [
          ColumnConfig False False 5 1 25 1
        , ColumnConfig False False 256 6 10 1
        ]
      , wsMergeCells         = [Merge ((1,1),(3,1))]
      , wsPageSetupPrs       = [PageSetupPr 1]
      , wsSheetViews         = [SheetView 1 1 0]
      , wsSheetFormat        = SheetFormat
        { sfCustomHeight     = True
        , sfDefaultColWidth  = 8
        , sfDefaultRowHeight = 12.75
        , sfOutlineLevelCol  = 0
        , sfOutlineLevelRow  = 0
        -- , sfX14Descent       = 0.25
        }
      , wsPageMargins = PageMargins 1 0.25 0.25 1 1 1
      , wsPageSetup = PageSetup 1 1 1 OrPortrait PODownThenOver 100 False
      }


