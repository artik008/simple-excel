{-# LANGUAGE OverloadedStrings #-}
module Defaults where

import           Data.ByteString.Lazy (ByteString)

import           Instances            ()
import           Models

relSharedStrings :: Relationship
relSharedStrings = Relationship
  { relID = 1
  , relTarget = "sharedStrings.xml"
  , relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  }

relStyles :: Relationship
relStyles = Relationship
  { relID = 2
  , relTarget = "styles.xml"
  , relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  }

relTheme :: Relationship
relTheme = Relationship
  { relID = 3
  , relTarget = "theme/theme1.xml"
  , relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
  }

orWorkbook :: Override
orWorkbook = Override
  { orContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
  , orPartName="/xl/workbook.xml"
  }

orRelsWorkbook :: Override
orRelsWorkbook = Override
  { orContentType="application/vnd.openxmlformats-package.relationships+xml"
  , orPartName="/xl/_rels/workbook.xml.rels"
  }

orSharedStrings :: Override
orSharedStrings = Override
  { orContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
  , orPartName="/xl/sharedStrings.xml"
  }

orStyles :: Override
orStyles = Override
  { orContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
  , orPartName="/xl/styles.xml"
  }

orCore :: Override
orCore = Override
  { orContentType="application/vnd.openxmlformats-package.core-properties+xml"
  , orPartName="/docProps/core.xml"
  }

orApp :: Override
orApp = Override
  { orContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"
  , orPartName="/docProps/app.xml"
  }

orRels :: Override
orRels = Override
  { orContentType="application/vnd.openxmlformats-package.relationships+xml"
  , orPartName="/_rels/.rels"
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

defStyle :: CellXf
defStyle = CellXf
  { border    = Just defBorder
  , fill      = Nothing
  , font      = Just $ Font "Helvetica Neue" False False 12 "000000" VCenter
  , numFmt    = Nothing
  , alignment = Nothing
  }

defTopStyle :: CellXf
defTopStyle = CellXf
  { border    = Just defTopBorder
  , fill      = Nothing
  , font      = Just $ Font "Helvetica Neue" False False 12 "000000" VCenter
  , numFmt    = Nothing
  , alignment = Nothing
  }

defBottomStyle :: CellXf
defBottomStyle = CellXf
  { border    = Just defBottomBorder
  , fill      = Nothing
  , font      = Just $ Font "Helvetica Neue" False False 12 "000000" VCenter
  , numFmt    = Nothing
  , alignment = Nothing
  }

defLeftStyle :: CellXf
defLeftStyle = CellXf
  { border    = Just defLeftBorder
  , fill      = Nothing
  , font      = Just $ Font "Helvetica Neue" False False 12 "000000" VCenter
  , numFmt    = Nothing
  , alignment = Nothing
  }

defRightStyle :: CellXf
defRightStyle = CellXf
  { border    = Just defRightBorder
  , fill      = Nothing
  , font      = Just $ Font "Helvetica Neue" False False 12 "000000" VCenter
  , numFmt    = Nothing
  , alignment = Nothing
  }

defBorder :: Border
defBorder = Border
  { borderStyles =
     [
       Just $ BorderStyle BLeft Thin (RGB "ffffff")
     , Just $ BorderStyle BRight Thin (RGB "ffffff")
     , Just $ BorderStyle BTop Thin (RGB "ffffff")
     , Just $ BorderStyle BBottom Thin (RGB "ffffff")
     , Just $ BorderStyle BDiagonal Thin (RGB "ffffff")
     ]
  }

defTopBorder :: Border
defTopBorder = Border
  { borderStyles =
     [
       Just $ BorderStyle BLeft Thin (RGB "ffffff")
     , Just $ BorderStyle BRight Thin (RGB "ffffff")
     , Just $ BorderStyle BTop Thin (RGB "ffffff")
     , Just $ BorderStyle BBottom Thin (RGB "808080")
     , Just $ BorderStyle BDiagonal Thin (RGB "ffffff")
     ]
  }

defBottomBorder :: Border
defBottomBorder = Border
  { borderStyles =
     [
       Just $ BorderStyle BLeft Thin (RGB "ffffff")
     , Just $ BorderStyle BRight Thin (RGB "ffffff")
     , Just $ BorderStyle BTop Thin (RGB "808080")
     , Just $ BorderStyle BBottom Thin (RGB "ffffff")
     , Just $ BorderStyle BDiagonal Thin (RGB "ffffff")
     ]
  }

defLeftBorder :: Border
defLeftBorder = Border
  { borderStyles =
     [
       Just $ BorderStyle BLeft Thin (RGB "ffffff")
     , Just $ BorderStyle BRight Thin (RGB "808080")
     , Just $ BorderStyle BTop Thin (RGB "ffffff")
     , Just $ BorderStyle BBottom Thin (RGB "ffffff")
     , Just $ BorderStyle BDiagonal Thin (RGB "ffffff")
     ]
  }

defRightBorder :: Border
defRightBorder = Border
  { borderStyles =
     [
       Just $ BorderStyle BLeft Thin (RGB "808080")
     , Just $ BorderStyle BRight Thin (RGB "ffffff")
     , Just $ BorderStyle BTop Thin (RGB "ffffff")
     , Just $ BorderStyle BBottom Thin (RGB "ffffff")
     , Just $ BorderStyle BDiagonal Thin (RGB "ffffff")
     ]
  }

defRowConfig :: RowConfig
defRowConfig = RowConfig False True False False 20 1

fileHeader :: ByteString
fileHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
