{-# LANGUAGE OverloadedStrings #-}
module Models where

import qualified Data.ByteString.Lazy    as BSL
import           Data.Char               (toLower)
import           Data.List               (intercalate)
import           Data.Map                (Map)
import           Data.String.Conversions (cs)
import           Data.Text               (Text)


type StyleID = Int

type CellCoords = (Int, Int)

data Merge = Merge Range deriving (Show, Eq)

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
  deriving (Show, Eq)

-- | Worksheets

data Worksheet = Worksheet
  { wsName           :: String
  , wsStyleFix       :: Bool
  -- ^ Style fix for Mac OS (set default style for empty cells in small tables)
  , wsCells          :: Map CellCoords Cell
  , wsRowsConfig     :: Map Int RowConfig
  , wsColumnsConfigs :: [ColumnConfig]
  , wsMergeCells     :: [Merge]
  , wsPageSetupPrs   :: [PageSetupPr]
  , wsSheetViews     :: [SheetView]
  , wsSheetFormat    :: SheetFormat
  , wsPageMargins    :: PageMargins
  , wsPageSetup      :: PageSetup
  }
  deriving (Show, Eq)

data IndexedWorksheet = IndexedWorksheet
  { iwsName           :: String
  , iwsDimension      :: Range
  , iwsRows           :: [IndRow]
  , iwsColumnsConfigs :: [ColumnConfig]
  , iwsMergeCells     :: [Merge]
  , iwsPageSetupPrs   :: [PageSetupPr]
  , iwsSheetViews     :: [SheetView]
  , iwsSheetFormat    :: SheetFormat
  , iwsPageMargins    :: PageMargins
  , iwsPageSetup      :: PageSetup
  }
  deriving (Show, Eq)

data IndRow = IndRow
  { irowNumber :: Int
  , irowConfig :: RowConfig
  , irowCells  :: [IndCell]
  }
  deriving (Show, Eq)

data PageSetupPr = PageSetupPr
  { psprFitToPage :: Int
  }
  deriving (Show, Eq)

data SheetView = SheetView
  { svDefaultGridColor :: Int
  , svShowGridLines    :: Int
  , svWorkbookViewId   :: Int
  }
  deriving (Show, Eq)

data SheetFormat = SheetFormat
  { sfCustomHeight     :: Bool
  , sfDefaultColWidth  :: Float
  , sfDefaultRowHeight :: Float
  , sfOutlineLevelCol  :: Int
  , sfOutlineLevelRow  :: Int
  -- , sfX14Descent       :: Float
  }
  deriving (Show, Eq)

data PageMargins = PageMargins
  { pmBottom :: Float
  , pmFooter :: Float
  , pmHeader :: Float
  , pmLeft   :: Float
  , pmRight  :: Float
  , pmTop    :: Float
  } deriving (Show, Eq)

data PageSetup = PageSetup
  { psFirstPageNumber    :: Int
  , psFitToHeight        :: Int
  , psFitToWidth         :: Int
  , psOrientation        :: Orientation
  , psPageOrder          :: PageOrder
  , psScale              :: Int
  , psUseFirstPageNumber :: Bool
  } deriving (Show, Eq)

data PageOrder = PODownThenOver | POOverThenDown deriving (Eq)

instance Show PageOrder where
  show PODownThenOver = "downThenOver"
  show POOverThenDown = "overThenDown"

data Orientation = OrDefault | OrPortrait | OrLandscape deriving (Eq)

instance Show Orientation where
  show OrDefault   = "default"
  show OrPortrait  = "portrait"
  show OrLandscape = "landscape"

data AppVariant = AppVariant
  { variantType  :: VariantType
  , variantValue :: Text
  }
  deriving (Show)

data VariantType =
    VarVariant
  | VarArray
  | VarVector
  | VarBlob
  | VarOblob
  | VarEmpty
  | VarNull
  | VarI1
  | VarI2
  | VarI4
  | VarI8
  | VarLpstr
  -- FIXME add all types
  deriving (Eq, Show)



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
  { fillType    :: Text
  , fillFgColor :: Color
  , fillBgColor :: Color
  }
  deriving (Show, Eq, Ord)

data Font = Font
  { fontName          :: Text
  , fontBold          :: Bool
  , fontOutline       :: Bool
  , fontSize          :: Int
  , fontColor         :: Color
  , fontVertAlignment :: FontVertAlType
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
data FontVertAlType = Baseline | Superscript | Subscript deriving (Eq, Ord)

instance Show HorAlType where
  show HCenter = "center"
  show HRight  = "right"
  show HLeft   = "left"

instance Show VertAlType where
  show VCenter = "center"
  show Top     = "top"
  show Bottom  = "bottom"

instance Show FontVertAlType where
  show Baseline    = "baseline"
  show Superscript = "superscript"
  show Subscript   = "subscript"

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
  deriving (Show, Eq)

data IndCell = IndCell
  { icellValue  :: Maybe IndCellValue
  , icellStyle  :: Int
  , icellCoords :: CellCoords
  }
  deriving (Show, Eq)

data IndCellValue
  = IndCellText Int
  | IndCellFloat Float
  | IndCellFormula Text (Maybe Text)
  deriving (Show, Eq)

data CellValue
  = CellText Text
  | CellFloat Float
  | CellFormula Text (Maybe Text)
  deriving (Show, Eq)

-- | Rows and Columns properties

data RowConfig = RowConfig
  { rowCollapsed    :: Bool
  , rowCustomFmt    :: Bool
  , rowCustomHeight :: Bool
  , rowHidden       :: Bool
  , rowHeight       :: Int
  , rowOutlineLevel :: Int
  }
  deriving (Show, Eq)


data ColumnConfig = ColumnConfig
  { columnBestFit     :: Bool
  , columnCustomWidth :: Bool
  , columnLastColumn  :: Int
  , columnFirstColumn :: Int
  , columnWidth       :: Float
  , columnStyle       :: Int
  }
  deriving (Show, Eq)


-- | Additional files

data Relationship = Relationship
  { relID     :: Int
  , relTarget :: FilePath
  , relType   :: URL
  }
  deriving (Show, Eq)

data Override = Override
  { orContentType :: String
  , orPartName    :: FilePath
  }
  deriving (Show, Eq)

-- | XMLTag

data XMLTag
  = XMLTag TagBegin TagBody TagEnd
  | XMLFloat Float
  | XMLInt Int
  | XMLText Text
  | XMLBool Bool
  deriving (Show, Eq)

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

