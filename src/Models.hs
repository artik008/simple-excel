{-# LANGUAGE OverloadedStrings #-}
module Models where

import qualified Data.ByteString.Lazy    as BSL
import           Data.List               (intercalate)
import           Data.Map                (Map)
import           Data.String.Conversions (cs)
import           Data.Text               (Text)


type StyleID = Int

type CellCoords = (Int, Int)

data Merge = Merge Range deriving (Show)

type Range = (CellCoords, CellCoords)

type URL = String

type TagBegin = (String, [Attr])
type TagBody = [XMLTag]
type TagEnd = Bool

type Attr = (String, String)

type CellName = String

-- ** Document

-- | XLSX

data Xlsx = Xlsx
  { worksheets :: [Worksheet]  -- FIXME map of worksheets
  }
  deriving (Show)

-- | Worksheets

data Worksheet = Worksheet
  { wsName           :: String
  , wsStyleFix       :: Bool
  -- ^ Style fix for Mac OS (set default style for empty cells in small tables)
  , wsCells          :: Map CellCoords Cell
  , wsRowsConfig     :: Map Int RowConfig
  , wsColumnsConfigs :: [ColumnConfig]
  , wsMergeCells     :: [Merge]
  , wsPageSetups     :: [PageSetup]
  , wsSheetViews     :: [SheetView]
  , wsSheetFormat    :: SheetFormat
  }
  deriving (Show)

data IndexedWorksheet = IndexedWorksheet
  { iwsName           :: String
  , iwsDimension      :: Range
  , iwsRows           :: [IndRow]
  , iwsColumnsConfigs :: [ColumnConfig]
  , iwsMergeCells     :: [Merge]
  , iwsPageSetups     :: [PageSetup]
  , iwsSheetViews     :: [SheetView]
  , iwsSheetFormat    :: SheetFormat
  }
  deriving (Show)

data IndRow = IndRow
  { irowNumber :: Int
  , irowConfig :: RowConfig
  , irowCells  :: [IndCell]
  }
  deriving (Show)

data PageSetup = PageSetup
  { psFitToPage :: Int
  }
  deriving (Show)

data SheetView = SheetView
  { svDefaultGridColor :: Int
  , svShowGridLines    :: Int
  , svWorkbookViewId   :: Int
  }
  deriving (Show)

data SheetFormat = SheetFormat
  { sfCustomHeight     :: Float
  , sfDefaultColWidth  :: Float
  , sfDefaultRowHeight :: Float
  , sfOutlineLevelCol  :: Int
  , sfOutlineLevelRow  :: Int
  }
  deriving (Show)

-- | Styles

data CellXf = CellXf
  { border    :: Maybe Border
  , fill      :: Maybe Fill
  , font      :: Maybe Font
  , numFmt    :: Maybe NumFmt
  , alignment :: Maybe Alignment
  }
  deriving (Show, Eq, Ord)

data IndCellXf = IndCellXf
  { icellXfIndex :: StyleID
  , iborder      :: Maybe StyleID
  , ifill        :: Maybe StyleID
  , ifont        :: Maybe StyleID
  , inumFmt      :: Maybe StyleID
  , ialignment   :: Maybe Alignment
  }
  deriving (Show, Eq, Ord)

data Border = Border
  { borderStyles :: [Maybe BorderStyle] }
  deriving (Show, Eq, Ord)


data BorderStyle = BorderStyle
  { borderSide  :: BorderSide
  , borderLine  :: BorderLine
  , borderColor :: Color
  }
  deriving (Show, Eq, Ord)

data BorderSide = BRight | BLeft | BTop | BBottom | BDiagonal
  deriving (Eq, Ord)

data BorderLine = Thin | Bold deriving (Show, Eq, Ord) -- FIXME

data Color = RGB Text | Indexed Int deriving (Eq, Ord) -- FIXME add indexed color

instance Show BorderSide where
  show BRight    = "right"
  show BLeft     = "left"
  show BTop      = "top"
  show BBottom   = "bottom"
  show BDiagonal = "diagonal"

instance Show Color where
  show (RGB c)     = cs c
  show (Indexed c) = show c

data Fill = Fill
  { fillType  :: Text
  , fillColor :: Color
  }
  deriving (Show, Eq, Ord)

data Font = Font
  { fontName          :: Text
  , fontBold          :: Bool
  , fontOutline       :: Bool
  , fontSize          :: Int
  , fontColor         :: Color
  , fontVertAlignment :: VertAlType
  }
  deriving (Show, Eq, Ord)

data NumFmt = NumFmt
  { formatCode :: String
  , numFmtId   :: StyleID
  }
  deriving (Show, Eq, Ord)

data Alignment = Alignment
  { horizontal :: HorAlType
  , vertical   :: VertAlType
  , wrapText   :: Text
  }
  deriving (Show, Eq, Ord)

data HorAlType = HCenter | HRight | HLeft deriving (Eq, Ord)
data VertAlType = VCenter | Top | Bottom deriving (Eq, Ord)

instance Show HorAlType where
  show HCenter = "center"
  show HRight  = "right"
  show HLeft   = "left"

instance Show VertAlType where
  show VCenter = "center"
  show Top     = "top"
  show Bottom  = "bottom"

data StyleSheet = StyleSheet
  { ssBorders :: Map Border Int
  , ssFills   :: Map Fill Int
  , ssFonts   :: Map Font Int
  , ssNumFmts :: Map NumFmt Int
  , ssCellXfs :: Map Int IndCellXf
  }
  deriving (Show, Eq, Ord)


-- | Cells

data Cell = Cell
  { cellValue :: Maybe CellValue
  , cellStyle :: CellXf
  }
  deriving (Show)

data IndCell = IndCell
  { icellValue  :: Maybe IndCellValue
  , icellStyle  :: Int
  , icellCoords :: CellCoords
  }
  deriving (Show)

data IndCellValue
  = IndCellText Int
  | IndCellFloat Float
  | IndCellFormula Text (Maybe Text)
  deriving (Show)

data CellValue
  = CellText Text
  | CellFloat Float
  | CellFormula Text (Maybe Text)
  deriving (Show)

-- | Rows and Columns properties

data RowConfig = RowConfig
  { rowCollapsed    :: Bool
  , rowCustomFmt    :: Bool
  , rowCustomHeight :: Bool
  , rowHidden       :: Bool
  , rowHeight       :: Int
  , rowOutlineLevel :: Int
  }
  deriving (Show)


data ColumnConfig = ColumnConfig
  { columnBestFit     :: Int
  , columnCustomWidth :: Int
  , columnLastColumn  :: Int
  , columnFirstColumn :: Int
  , columnWidth       :: Int
  }
  deriving (Show)


-- | Additional files

data Relationship = Relationship
  { relID     :: Int
  , relTarget :: FilePath
  , relType   :: URL
  }
  deriving (Show)

data Override = Override
  { orContentType :: String
  , orPartName    :: FilePath
  }
  deriving (Show)

-- | XMLTag

data XMLTag
  = XMLTag TagBegin TagBody TagEnd
  | XMLFloat Float
  | XMLInt Int
  | XMLText Text
  deriving (Show)

-- | Indexes history

type StyleHistory = Map CellXf Int
type StringHistory = Map Text Int

data WithHistory a = WithHistory [a] StyleHistory StringHistory


-- | ToExcelXML - class for types for transforming to OpenXML

class (Show a) => ToExcelXML a where
  toTag :: a -> XMLTag
  toTagList :: a -> [XMLTag]
  toTagList v = [toTag v]
  (<++) :: XMLTag -> a -> XMLTag
  (<++) (XMLTag (name, args) body end) element =
    XMLTag (name,args) (toTag element:body) True
  (<<<) :: XMLTag -> [a] -> XMLTag
  (<<<) tag l = foldl (<++) tag l

(-++) :: BSL.ByteString -> XMLTag -> BSL.ByteString
(-++) bsl tag = bsl <> toXML tag

toXML :: XMLTag -> BSL.ByteString
toXML (XMLTag (name, args) body end) =
  (
    cs $
      "<" <> name <> "\n" <>
      (intercalate "\n" $
          map (\(x,y) -> x <> "=\"" <> y <> "\"") args)
  ) <>
  if end
    then ">" <>
      BSL.intercalate "\n" (map toXML $ reverse body) <>
      "\n</" <> cs name <> ">"
    else "/>"
toXML (XMLFloat v) = cs $ show v
toXML (XMLInt v) = cs $ show v
toXML (XMLText v) = cs v

