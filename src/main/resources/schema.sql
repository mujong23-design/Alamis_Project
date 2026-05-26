-- 사용자
CREATE TABLE IF NOT EXISTS users (
    user_id     VARCHAR(50)  PRIMARY KEY,
    user_pw     VARCHAR(100) NOT NULL,
    user_name   VARCHAR(50)  NOT NULL,
    role        VARCHAR(20)  DEFAULT 'USER',
    created_at  TIMESTAMP    DEFAULT CURRENT_TIMESTAMP
);

-- 입고 (1 입고 = 1 row, outbound_qty 누적으로 출고 관리)
CREATE TABLE IF NOT EXISTS inbound (
    inbound_id      VARCHAR(50)   PRIMARY KEY,
    production_date DATE          NOT NULL,
    package_unit_kg NUMERIC(8,3)  NOT NULL CHECK (package_unit_kg > 0),
    package_count   INTEGER       NOT NULL CHECK (package_count > 0),
    size_um         NUMERIC(5,3),
    outbound_qty    INTEGER       NOT NULL DEFAULT 0 CHECK (outbound_qty >= 0),
    printed_yn      CHAR(1)       NOT NULL DEFAULT 'N' CHECK (printed_yn IN ('Y','N')),
    remark          VARCHAR(500),
    created_at      TIMESTAMP     DEFAULT CURRENT_TIMESTAMP,
    CHECK (outbound_qty <= package_count)
);

CREATE INDEX IF NOT EXISTS idx_inbound_date ON inbound(production_date);

-- 출고 마스터 (1 outbound = N inbound items)
CREATE TABLE IF NOT EXISTS outbound (
    outbound_id        VARCHAR(50)  PRIMARY KEY,     -- O-YYYYMMDD-NNN
    outbound_date      DATE         NOT NULL,
    coa_code           VARCHAR(100) NOT NULL,         -- 성적서 코드 (필수)
    outbound_location  VARCHAR(200),
    remark             VARCHAR(500),
    created_at         TIMESTAMP    DEFAULT CURRENT_TIMESTAMP
);

-- 출고 디테일
CREATE TABLE IF NOT EXISTS outbound_item (
    item_id      BIGSERIAL    PRIMARY KEY,
    outbound_id  VARCHAR(50)  NOT NULL REFERENCES outbound(outbound_id) ON DELETE CASCADE,
    inbound_id   VARCHAR(50)  NOT NULL REFERENCES inbound(inbound_id),
    quantity     INTEGER      NOT NULL CHECK (quantity > 0),
    created_at   TIMESTAMP    DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_outbound_date          ON outbound(outbound_date);
CREATE INDEX IF NOT EXISTS idx_outbound_item_outbound ON outbound_item(outbound_id);
CREATE INDEX IF NOT EXISTS idx_outbound_item_inbound  ON outbound_item(inbound_id);
