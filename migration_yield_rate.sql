-- 出成率 / 熟料單價 功能
-- 出成率預設 1（不影響現有資料與計算），熟料單價由 DB 自動算出，不可手動寫入
ALTER TABLE cost_materials ADD COLUMN IF NOT EXISTS yield_rate NUMERIC NOT NULL DEFAULT 1;
ALTER TABLE cost_materials ADD COLUMN IF NOT EXISTS cooked_unit_price NUMERIC GENERATED ALWAYS AS (unit_cost / NULLIF(yield_rate, 0)) STORED;
