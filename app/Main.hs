{-# LANGUAGE OverloadedStrings #-}
module Main where

-- import qualified Codec.Archive.Zip     as Zip
-- import           Control.Lens
-- import qualified Data.ByteString.Lazy  as BSL
-- import qualified Data.Map              as Map
-- import           Data.Time.Clock.POSIX
-- import           System.Directory

-- import           Codec.Xlsx

main :: IO ()
main = putStrLn "Empty" -- zipXlsx "report"


-- testXlsx :: IO ()
-- testXlsx = do
--   ct <- getPOSIXTime
--   let
--       sheet = def
--                   & cellValueAt (1,2) ?~ CellDouble 42.0
--                   & cellValueAt (3,2) ?~ CellText "foo"
--       xlsx = def & atSheet "List1" ?~ sheet
--         { _wsColumnsProperties = cp
--         , _wsRowPropertiesMap = rp
--         }
--   BSL.writeFile "example.xlsx" $ fromXlsx ct xlsx

-- cp :: [ColumnsProperties]
-- cp = [ ColumnsProperties 1 2 Nothing Nothing False True True
--      , ColumnsProperties 1 2 Nothing Nothing False True True
--      ]

-- rp :: Map.Map Int RowProperties
-- rp = foldr (\(x,y) l -> Map.insert x y l) Map.empty
--   [ (1, deffrp)
--   , (2, deffrp)
--   , (3, deffrp)
--   ]
--   where
--     deffrp = RowProps (Just (AutomaticHeight 5)) Nothing False


-- zipXlsx :: FilePath -> IO ()
-- zipXlsx fname = do
--   files <- filesInDir fname
--   openedFiles <- sequence $
--     map
--       (\x -> do
--         f <- BSL.readFile x
--         return $ (drop (length fname + 1) x, f)
--       )
--       files
--   BSL.writeFile "test.xlsx" $ Zip.fromArchive $
--     foldr
--       Zip.addEntryToArchive
--       Zip.emptyArchive
--       (map (\(x, y) -> Zip.toEntry x 5 y) openedFiles)
--   where
--     filesInDir name = do
--       inDir <- listDirectory name
--       checked <- sequence $
--         map (\f -> do
--           isDir <- doesDirectoryExist (name <> "/" <> f)
--           if isDir
--           then filesInDir (name <> "/" <> f)
--           else return [name <> "/" <> f]
--         ) inDir
--       return $ concat checked













