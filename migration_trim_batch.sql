-- 修清率 + 批次總重覆寫
-- 1) 食材新增修清率 trim_rate（預設 1，不影響現有計算）
ALTER TABLE cost_materials ADD COLUMN IF NOT EXISTS trim_rate NUMERIC NOT NULL DEFAULT 1;

-- 2) 熟料單價 = 單位成本 ÷ (修清率 × 熟成率)
ALTER TABLE cost_materials DROP COLUMN IF EXISTS cooked_unit_price;
ALTER TABLE cost_materials ADD COLUMN cooked_unit_price NUMERIC
  GENERATED ALWAYS AS (unit_cost / NULLIF(trim_rate * yield_rate, 0)) STORED;

-- 3) 半成品批次總重（留空＝配方用量加總；有填則以此為產出量）
ALTER TABLE cost_semis ADD COLUMN IF NOT EXISTS batch_weight_g NUMERIC;
