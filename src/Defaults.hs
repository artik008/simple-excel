{-# LANGUAGE OverloadedStrings #-}
module Defaults where

import           Data.ByteString.Lazy (ByteString)
import           Data.Text            (Text, pack)

import           Instances            ()
import           Models
import           Utils

relSharedStrings :: Relationship
relSharedStrings = Relationship
  { relID = 1
  , relTarget = "sharedStrings.xml"
  , relType = urlHeader <> "officeDocument/2006/relationships/sharedStrings"
  }

relStyles :: Relationship
relStyles = Relationship
  { relID = 2
  , relTarget = "styles.xml"
  , relType = urlHeader <> "officeDocument/2006/relationships/styles"
  }

relWorkBook :: Relationship
relWorkBook = Relationship
  { relID = 1
  , relTarget = "xl/workbook.xml"
  , relType = urlHeader <> "officeDocument/2006/relationships/officeDocument"
  }

relCore :: Relationship
relCore = Relationship
  { relID = 2
  , relTarget = "docProps/core.xml"
  , relType =
      urlHeader <> "officeDocument/2006/relationships/metadata/core-properties"
  }

relApp :: Relationship
relApp = Relationship
  { relID = 3
  , relTarget = "docProps/app.xml"
  , relType =
      urlHeader <> "officeDocument/2006/relationships/extended-properties"
  }

orWorkbook :: Override
orWorkbook = Override
  { orContentType =
      appPathHead <> "officedocument.spreadsheetml.sheet.main+xml"
  , orPartName = "/xl/workbook.xml"
  }

orRelsWorkbook :: Override
orRelsWorkbook = Override
  { orContentType = appPathHead <> "package.relationships+xml"
  , orPartName = "/xl/_rels/workbook.xml.rels"
  }

orSharedStrings :: Override
orSharedStrings = Override
  { orContentType =
      appPathHead <> "officedocument.spreadsheetml.sharedStrings+xml"
  , orPartName = "/xl/sharedStrings.xml"
  }

orStyles :: Override
orStyles = Override
  { orContentType = appPathHead <> "officedocument.spreadsheetml.styles+xml"
  , orPartName = "/xl/styles.xml"
  }

orCore :: Override
orCore = Override
  { orContentType = appPathHead <> "package.core-properties+xml"
  , orPartName = "/docProps/core.xml"
  }

orApp :: Override
orApp = Override
  { orContentType = appPathHead <> "officedocument.extended-properties+xml"
  , orPartName = "/docProps/app.xml"
  }

orRels :: Override
orRels = Override
  { orContentType = appPathHead <> "package.relationships+xml"
  , orPartName = "/_rels/.rels"
  }

overrides :: [Override]
overrides =
  [ orWorkbook
  , orRelsWorkbook
  , orSharedStrings
  , orStyles
  , orCore
  , orApp
  , orRels
  ]

appHeadingPairs :: Int -> [AppVariant]
appHeadingPairs i =
  [ AppVariant VarLpstr "Листы"
  , AppVariant VarI4 (pack $ show i)
  ]

appTitlesOfParts :: [Text] -> [XMLTag]
appTitlesOfParts wsNames = map (\x -> valueTag "vt:lpstr" x) wsNames


emptyStyle :: CellXf
emptyStyle = CellXf
  { border    = Nothing
  , fill      = Nothing
  , font      = Nothing
  , numFmt    = Nothing
  , alignment = Nothing
  }


defStyle :: CellXf
defStyle = CellXf
  { border    = Just defBorder
  , fill      = Nothing
  , font      = Just $ Font "Times New Roman" False False 12 "000000" Baseline
  , numFmt    = Nothing
  , alignment = Nothing
  }

defOutStyle :: CellXf
defOutStyle = defStyle { border = Just defOutBorder }

defTopStyle :: CellXf
defTopStyle = defStyle { border = Just defTopBorder }

defBottomStyle :: CellXf
defBottomStyle = defStyle { border = Just defBottomBorder }

defLeftStyle :: CellXf
defLeftStyle = defStyle { border = Just defLeftBorder }

defRightStyle :: CellXf
defRightStyle = defStyle { border = Just defRightBorder }

defBorder :: Border
defBorder = Border
  { borderStyles = map borderStyleGray
     [BLeft, BRight, BTop, BBottom, BDiagonal]
  }

defOutBorder :: Border
defOutBorder = Border
  { borderStyles = map borderStyleWhite
     [BLeft, BRight, BTop, BBottom, BDiagonal]
  }

defTopBorder :: Border
defTopBorder = Border
  { borderStyles = oneSideGray BBottom
  }

defBottomBorder :: Border
defBottomBorder = Border
  { borderStyles = oneSideGray BTop
  }

defLeftBorder :: Border
defLeftBorder = Border
  { borderStyles = oneSideGray BRight
  }

defRightBorder :: Border
defRightBorder = Border
  { borderStyles = oneSideGray BLeft
  }

oneSideGray :: BorderSide -> [Maybe BorderStyle]
oneSideGray bSide =
     (
      borderStyleGray bSide
      :
      map borderStyleWhite (filter (/=bSide) allSides)
     )

allSides :: [BorderSide]
allSides = [BLeft, BRight, BTop, BBottom, BDiagonal]

borderStyleGray :: BorderSide -> Maybe BorderStyle
borderStyleGray bSide = Just $ BorderStyle bSide Thin (RGB "808080")

borderStyleWhite :: BorderSide -> Maybe BorderStyle
borderStyleWhite bSide = Just $ BorderStyle bSide Thin (RGB "ffffff")

defRowConfig :: RowConfig
defRowConfig = RowConfig False True False False 20 1

fileHeader :: ByteString
fileHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\"  standalone=\"yes\"?>"

urlHeader :: URL
urlHeader = "http://schemas.openxmlformats.org/"

appPathHead :: FilePath
appPathHead = "application/vnd.openxmlformats-"
