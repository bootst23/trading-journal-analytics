-- ============================================================
-- Trading Journal Analysis (SQLite) - R-multiple based
-- Author: Nathan Yap Hao Kang
--
-- Setup (typed table first - a raw .import makes every column
-- TEXT and silently breaks numeric comparisons):
--
--   sqlite3 journal.db
--   CREATE TABLE trades (
--       date TEXT, month TEXT, direction TEXT, trade_no INTEGER,
--       result TEXT, r REAL, v1 TEXT, v2 TEXT, v3 TEXT, v4 TEXT,
--       violations INTEGER, type TEXT, grade TEXT, mins_in_trade REAL);
--   .import --csv --skip 1 nathan_journal.csv trades
--   .mode column
--   .headers on
-- ============================================================

-- Shared equity curve: one row per trade in sequence (date, then trade_no).
-- The peak starts at 0R, so early losses count as drawdown (as in dashboard.html).
DROP VIEW IF EXISTS equity_curve;
CREATE VIEW equity_curve AS
WITH running AS (
    SELECT date, trade_no, r, SUM(r) OVER (ORDER BY date, trade_no) AS cum_r
    FROM trades
),
peaks AS (
    SELECT *, MAX(0, MAX(cum_r) OVER (ORDER BY date, trade_no)) AS peak_r
    FROM running
)
SELECT date, trade_no, r, cum_r, peak_r, cum_r - peak_r AS drawdown_r
FROM peaks;

-- 1. Headline statistics
SELECT
    COUNT(*)                                              AS total_trades,
    SUM(result = 'Win')                                   AS wins,
    SUM(result = 'Loss')                                  AS losses,
    SUM(result = 'BE')                                    AS breakeven,
    ROUND(100.0 * SUM(result = 'Win')
        / (SUM(result = 'Win') + SUM(result = 'Loss')), 1) AS win_rate_excl_be,
    ROUND(SUM(r), 2)                                      AS net_r,
    ROUND(AVG(r), 3)                                      AS expectancy_r
FROM trades;

-- 2. Payoff ratio: average win vs average loss (in R)
SELECT
    ROUND(AVG(CASE WHEN result = 'Win'  THEN r END), 2)  AS avg_win_r,
    ROUND(AVG(CASE WHEN result = 'Loss' THEN r END), 2)  AS avg_loss_r,
    ROUND(AVG(CASE WHEN result = 'Win' THEN r END)
        / ABS(AVG(CASE WHEN result = 'Loss' THEN r END)), 2) AS payoff_ratio
FROM trades;

-- 3. THE key finding: rule-compliant vs violation trades
SELECT
    CASE WHEN violations = 0 THEN 'rule-compliant' ELSE 'violated rules' END AS discipline,
    COUNT(*)               AS trades,
    ROUND(SUM(r), 2)       AS net_r,
    ROUND(AVG(r), 3)       AS expectancy_r
FROM trades
GROUP BY violations = 0
ORDER BY expectancy_r DESC;

-- 4. Expectancy by setup grade (bucketed A/B/C/D)
SELECT
    SUBSTR(grade, 1, 1)    AS grade_bucket,
    COUNT(*)               AS trades,
    ROUND(SUM(r), 2)       AS net_r,
    ROUND(AVG(r), 3)       AS expectancy_r,
    ROUND(100.0 * SUM(result = 'Win')
        / NULLIF(SUM(result = 'Win') + SUM(result = 'Loss'), 0), 1) AS win_rate
FROM trades
GROUP BY grade_bucket
ORDER BY grade_bucket;

-- 5. Grade discipline over time: share of trades that were A-grade, by month
SELECT
    strftime('%Y-%m', date)                             AS month,
    COUNT(*)                                            AS trades,
    ROUND(100.0 * SUM(SUBSTR(grade,1,1) = 'A') / COUNT(*), 1) AS pct_a_grade,
    ROUND(SUM(r), 2)                                    AS net_r
FROM trades
GROUP BY month
ORDER BY month;

-- 6. Performance by trade type
SELECT
    type,
    COUNT(*)          AS trades,
    ROUND(SUM(r), 2)  AS net_r,
    ROUND(AVG(r), 3)  AS expectancy_r
FROM trades
GROUP BY type
ORDER BY net_r DESC;

-- 7. Equity curve and drawdown in R
SELECT date, trade_no, ROUND(cum_r, 2) AS equity_r, ROUND(drawdown_r, 2) AS drawdown_r
FROM equity_curve ORDER BY date, trade_no;

-- 8. Maximum drawdown suffered (R)
SELECT ROUND(MIN(drawdown_r), 2) AS max_drawdown_r FROM equity_curve;

-- 9. Longest losing streak (gaps-and-islands)
WITH flagged AS (
    SELECT date, trade_no,
        CASE WHEN result = 'Loss' THEN 1 ELSE 0 END AS is_loss,
        ROW_NUMBER() OVER (ORDER BY date, trade_no)
      - ROW_NUMBER() OVER (PARTITION BY CASE WHEN result = 'Loss' THEN 1 ELSE 0 END
                           ORDER BY date, trade_no) AS grp
    FROM trades
)
SELECT MAX(streak) AS longest_losing_streak
FROM (SELECT COUNT(*) AS streak FROM flagged WHERE is_loss = 1 GROUP BY grp);

-- 10. Violation rate by month: did discipline improve?
SELECT
    strftime('%Y-%m', date)                                   AS month,
    COUNT(*)                                                  AS trades,
    ROUND(100.0 * SUM(violations > 0) / COUNT(*), 1)          AS violation_rate_pct,
    ROUND(SUM(r), 2)                                          AS net_r
FROM trades
GROUP BY month
ORDER BY month;
