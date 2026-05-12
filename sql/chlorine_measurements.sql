IF OBJECT_ID('dbo.Chlorine', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.Chlorine (
        id INT IDENTITY(1,1) PRIMARY KEY,
        measured_at DATETIME2 NOT NULL,
        chlorine_value DECIMAL(10, 2) NOT NULL,
        unit NVARCHAR(20) NOT NULL CONSTRAINT DF_Chlorine_unit DEFAULT 'mg/L',
        note NVARCHAR(1000) NULL,
        created_at DATETIME2 NOT NULL CONSTRAINT DF_Chlorine_created_at DEFAULT SYSUTCDATETIME()
    );
END;
