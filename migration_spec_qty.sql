-- 規格數量欄位：讓單位成本可以用「進貨單價 ÷ 規格數量」公式自動算，比照原始 Excel 的邏輯
-- 留空 = 沿用現有單位成本手動維護模式（完全向下相容，不影響任何現有資料/計算）
ALTER TABLE cost_materials ADD COLUMN IF NOT EXISTS spec_qty NUMERIC;
