-- PostgreSQL 数据验证测试目标表
-- 用于「数据验证排查」向导步骤①：数据库类型选 PostgreSQL，粘贴本 DDL 后点「解析 DDL」
CREATE TABLE IF NOT EXISTS t_data_valid_demo (
    id             INTEGER      NOT NULL,
    patient_uuid   UUID         NOT NULL,
    patient_name   VARCHAR(50)  NOT NULL,
    age            SMALLINT     NOT NULL,
    birth_date     DATE,
    created_at     TIMESTAMP,
    last_updated   TIMESTAMPTZ,
    start_time     TIME,
    zone_time      TIMETZ,
    is_active      BOOLEAN      NOT NULL,
    amount         NUMERIC(10,2),
    ratio          NUMERIC(8,4),
    record_count   BIGINT,
    ext_data       JSON,
    settings       JSONB,
    raw_tags       TEXT[],
    duration_interval INTERVAL,
    zjb_note       VARCHAR(100),
    dmmc           VARCHAR(100),
    PRIMARY KEY (id)
);