-- Newton Raphson Root Finding Method to find NPV = 0
-- Reviewed by: Ryan A I 
-- Date: 22/07/2024
--

-- The script uses the combination of Bisection method and Newton Raphson method
-- Bisection method is use for 5 iteration to have a better guess for newton raphson to start with
-- The newton rapshon is then run until max iterations pass within the function.
-- Or until the NPV is less than the tolerance

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER FUNCTION [dbo].[CalculateIRRBiNROpt]
(
    @CashFlows NVARCHAR(MAX),
    @MaxIterations INT = 20,
    @Tolerance FLOAT = 0.0000001
)
-- RETURNS VARCHAR(MAX)
RETURNS FLOAT
AS
BEGIN
    DECLARE @IRR FLOAT;
    DECLARE @PrevIRR FLOAT = 0;
    DECLARE @Iteration INT = 0;
    DECLARE @NPV NUMERIC(38, 6) = 0.0;
    DECLARE @NPVA NUMERIC(38, 6) = 0.0;
    DECLARE @NPVB NUMERIC(38, 6) = 0.0;
    DECLARE @NPVC NUMERIC(38, 6) = 0.0;
    DECLARE @A FLOAT = -0.9;
    DECLARE @B FLOAT = 1.0;
    DECLARE @C FLOAT = 0;
    DECLARE @DiffNPV NUMERIC(38, 6) = 0.0;
    DECLARE @DeltaIRR NUMERIC(38, 6) = 0.0;
    DECLARE @Converged BIT = 0;
    DECLARE @Result VARCHAR(MAX);
    DECLARE @CurrentIteration INT = 0;

    
    DECLARE @StartTime DATETIME2(7) = SYSDATETIME();
    DECLARE @EndTime DATETIME2(7);
    DECLARE @ExecutionTime VARCHAR(50);
    
    DECLARE @CashFlowTable TABLE (CashFlowAmount NUMERIC(38, 6), IndexNum INT IDENTITY(0,1));
    INSERT INTO @CashFlowTable
    SELECT CAST(value AS NUMERIC(38, 6)) 
    FROM STRING_SPLIT(@CashFlows, ',');

    
     -- Check Positive cashflow and Negative Cashflow --
    DECLARE @NegativeFlows NUMERIC(38, 6), @PositiveFlows NUMERIC(38, 6);
    SELECT 
        @NegativeFlows = ABS(SUM(CASE WHEN CashFlowAmount < 0 THEN CashFlowAmount ELSE 0 END)),
        @PositiveFlows = SUM(CASE WHEN CashFlowAmount > 0 THEN CashFlowAmount ELSE 0 END)
    FROM @CashFlowTable;

    IF (@NegativeFlows = 0 OR @PositiveFlows = 0)
    BEGIN
        SET @IRR = NULL; -- IRR is undefined in cases where there are no both positive and negative cashflow
        -- SET @Result = CONCAT(CAST(@C AS VARCHAR(50)),',',CAST(@IRR AS VARCHAR(50)),',',CAST(@Iteration AS varchar(50)),',',CAST(@Converged AS varchar(50)),',',@ExecutionTime);
        -- RETURN @Result;
        RETURN @IRR;
    END;

    -- Start Bisection Method
    WHILE @Iteration < 6
    BEGIN
        -- Calculate the midpoint (c) --
        SET @C = (@A + @B) / 2;

        -- Calculate NPV for IRR at point @A, @B, and @C --
        SELECT 
            @NPVA = SUM(CashFlowAmount / POWER(1 + @A, IndexNum)), -- NPV at IRR @A
            @NPVB = SUM(CashFlowAmount / POWER(1 + @B, IndexNum)), -- NPV at IRR @B
            @NPVC = SUM(CashFlowAmount / POWER(1 + @C, IndexNum))  -- NPV at IRR @C
        FROM @CashFlowTable;

        -- Find the position where root lies --
        IF @NPVA * @NPVC < 0
            SET @B = @C;  -- Root is between @A and @C
        ELSE
            SET @A = @C;  -- Root is between @C and @B

        SET @Iteration += 1;
    END;

    SET @CurrentIteration = @Iteration;
    SET @Iteration = 0;

    SET @IRR = @C;

    -- Solve the non-linear equation of IRR using Newton-Raphson method with initial guess calculated using bisection method --
    WHILE @Iteration < @MaxIterations
    BEGIN
        SELECT 
            @NPV = SUM(CashFlowAmount / NULLIF(POWER(1 + @IRR, IndexNum),0)), -- Calculate the NPV
            @DiffNPV = CAST(SUM(-IndexNum * CashFlowAmount / NULLIF(POWER(1 + @IRR, IndexNum + 1),0)) AS NUMERIC(38, 6)) -- Calculate the derivative of NPV
        FROM @CashFlowTable;

        -- Check the condition for the derivative of NPV, if it's null or it's smaller than the tolerance. Break if satisfied and set Converged to TRUE--
        IF @DiffNPV IS NULL OR ABS(@NPV) < @Tolerance 
        BEGIN
            SET @Converged = 1;
            BREAK;
        END;

        -- Part of Newton-Raphson algorithm, subtract the current IRR with delta IRR [ IRR = = IRR -f(x)/f'(x)]  --
        SET @DeltaIRR = @NPV / NULLIF(@DiffNPV, 0); 
        SET @IRR = @IRR - @DeltaIRR;

        -- Check if delta IRR exceed 10 
        IF ABS(@DeltaIRR) > 10
        BEGIN
            SET @Converged = 0;
            BREAK;
        END;

        -- Check if IRR is less than or equal to -1 set the converged to 0
        IF @IRR <= -1
        BEGIN
            SET @Converged = 0;
            BREAK;
        END;

        -- IF ABS(@PrevIRR - @IRR) <= @Tolerance
        -- BEGIN
        --     SET @Converged = 1;
        --     BREAK;
        -- END;

        -- -- Check if ratio of DeltaIRR divided by the current IRR is smaller than the tolerance break the calculation and Set converged to TRUE--
        -- IF @IRR IS NULL OR ABS(@DeltaIRR / NULLIF(@IRR, 0)) < @Tolerance
        -- BEGIN
        --     SET @Converged = 1;
        --     BREAK;
        -- END;

        -- Set the Prev IRR to current IRR, increase the iteration by 1 and loop through the process of Newton-Raphson --
        SET @PrevIRR = @IRR;
        SET @Iteration += 1;
    END

    IF @Iteration >= @MaxIterations OR @Converged = 0
            SET @IRR = NULL;

    SET @CurrentIteration += @Iteration;

    -- SET @EndTime = SYSDATETIME();
    -- DECLARE @ElapsedTime FLOAT = CAST(DATEDIFF_BIG(MICROSECOND, @StartTime, @EndTime) AS FLOAT) / 1000.0;
    -- SET @ExecutionTime = FORMAT(@ElapsedTime, '0.000000 ms');


    -- SET @Result = CONCAT(CAST(@C AS VARCHAR(50)),',',CAST(@IRR AS VARCHAR(50)),',',CAST(@Iteration AS varchar(50)),',',CAST(@Converged AS varchar(50)),',',@ExecutionTime);
    -- RETURN @Result;

    RETURN @IRR;
END;
GO
