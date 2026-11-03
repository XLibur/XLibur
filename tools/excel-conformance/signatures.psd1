@{
    # N name, F family, A argument kinds, T typical argument text per argument,
    # Min fewest arguments Excel accepts (default: all of A).
    # Kinds: num, int, text, bool, any, ref. Variadic functions list two arguments.
    # Left out on purpose: volatile or environment functions (TODAY, NOW, RAND, INFO, CELL),
    # functions that return arrays, and functions that need an _xlfn. prefix.
    Functions = @(
        # math
        @{ N = 'SIGN'; F = 'math'; A = @('num'); T = @('-5') }
        @{ N = 'SQRT'; F = 'math'; A = @('num'); T = @('16') }
        @{ N = 'EXP'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'LN'; F = 'math'; A = @('num'); T = @('10') }
        @{ N = 'LOG'; F = 'math'; A = @('num', 'num'); T = @('100', '10'); Min = 1 }
        @{ N = 'LOG10'; F = 'math'; A = @('num'); T = @('1000') }
        @{ N = 'CEILING'; F = 'math'; A = @('num', 'num'); T = @('2.5', '1') }
        @{ N = 'FLOOR'; F = 'math'; A = @('num', 'num'); T = @('2.5', '1') }
        @{ N = 'ROUNDUP'; F = 'math'; A = @('num', 'int'); T = @('2.345', '2') }
        @{ N = 'ROUNDDOWN'; F = 'math'; A = @('num', 'int'); T = @('2.345', '2') }
        @{ N = 'TRUNC'; F = 'math'; A = @('num', 'int'); T = @('8.9', '0'); Min = 1 }
        @{ N = 'FACT'; F = 'math'; A = @('int'); T = @('5') }
        @{ N = 'FACTDOUBLE'; F = 'math'; A = @('int'); T = @('7') }
        @{ N = 'COMBIN'; F = 'math'; A = @('int', 'int'); T = @('5', '2') }
        @{ N = 'GCD'; F = 'math'; A = @('int', 'int'); T = @('12', '18'); Min = 1 }
        @{ N = 'LCM'; F = 'math'; A = @('int', 'int'); T = @('4', '6'); Min = 1 }
        @{ N = 'QUOTIENT'; F = 'math'; A = @('num', 'num'); T = @('7', '2') }
        @{ N = 'EVEN'; F = 'math'; A = @('num'); T = @('3') }
        @{ N = 'ODD'; F = 'math'; A = @('num'); T = @('2') }
        @{ N = 'MROUND'; F = 'math'; A = @('num', 'num'); T = @('10', '3') }
        @{ N = 'PI'; F = 'math'; A = @(); T = @() }
        @{ N = 'SIN'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'COS'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'TAN'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'ASIN'; F = 'math'; A = @('num'); T = @('0.5') }
        @{ N = 'ACOS'; F = 'math'; A = @('num'); T = @('0.5') }
        @{ N = 'ATAN'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'ATAN2'; F = 'math'; A = @('num', 'num'); T = @('1', '1') }
        @{ N = 'DEGREES'; F = 'math'; A = @('num'); T = @('3.14159265358979') }
        @{ N = 'RADIANS'; F = 'math'; A = @('num'); T = @('180') }
        @{ N = 'ROMAN'; F = 'math'; A = @('int', 'any'); T = @('1999', '0'); Min = 1 }
        @{ N = 'MAX'; F = 'math'; A = @('any', 'any'); T = @('3', '7'); Min = 1 }
        @{ N = 'MIN'; F = 'math'; A = @('any', 'any'); T = @('3', '7'); Min = 1 }
        @{ N = 'SUMSQ'; F = 'math'; A = @('any', 'any'); T = @('3', '4'); Min = 1 }
        @{ N = 'SUMPRODUCT'; F = 'math'; A = @('ref', 'ref'); T = @('A1:A3', 'B1:B3'); Min = 1 }

        # text
        @{ N = 'CONCATENATE'; F = 'text'; A = @('text', 'text'); T = @('"a"', '"b"'); Min = 1 }
        @{ N = 'EXACT'; F = 'text'; A = @('text', 'text'); T = @('"a"', '"A"') }
        @{ N = 'LOWER'; F = 'text'; A = @('text'); T = @('"AbC"') }
        @{ N = 'UPPER'; F = 'text'; A = @('text'); T = @('"AbC"') }
        @{ N = 'PROPER'; F = 'text'; A = @('text'); T = @('"hello wORLD"') }
        @{ N = 'RIGHT'; F = 'text'; A = @('text', 'int'); T = @('"hello"', '2'); Min = 1 }
        @{ N = 'REPT'; F = 'text'; A = @('text', 'int'); T = @('"ab"', '3') }
        @{ N = 'REPLACE'; F = 'text'; A = @('text', 'int', 'int', 'text'); T = @('"abcdef"', '2', '3', '"X"') }
        @{ N = 'SEARCH'; F = 'text'; A = @('text', 'text', 'int'); T = @('"l"', '"hello"', '1'); Min = 2 }
        @{ N = 'CHAR'; F = 'text'; A = @('int'); T = @('65') }
        @{ N = 'CODE'; F = 'text'; A = @('text'); T = @('"A"') }
        @{ N = 'CLEAN'; F = 'text'; A = @('text'); T = @('"abc"') }
        @{ N = 'T'; F = 'text'; A = @('any'); T = @('"abc"') }

        # logical
        @{ N = 'NOT'; F = 'logical'; A = @('any'); T = @('TRUE') }
        @{ N = 'OR'; F = 'logical'; A = @('any', 'any'); T = @('FALSE', 'TRUE'); Min = 1 }
        @{ N = 'TRUE'; F = 'logical'; A = @(); T = @() }
        @{ N = 'FALSE'; F = 'logical'; A = @(); T = @() }

        # information
        @{ N = 'ISBLANK'; F = 'information'; A = @('any'); T = @('A1') }
        @{ N = 'ISERR'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISERROR'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISEVEN'; F = 'information'; A = @('num'); T = @('4') }
        @{ N = 'ISODD'; F = 'information'; A = @('num'); T = @('3') }
        @{ N = 'ISLOGICAL'; F = 'information'; A = @('any'); T = @('TRUE') }
        @{ N = 'ISNA'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISNONTEXT'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISTEXT'; F = 'information'; A = @('any'); T = @('"a"') }
        @{ N = 'ISREF'; F = 'information'; A = @('any'); T = @('A1') }
        @{ N = 'NA'; F = 'information'; A = @(); T = @() }
        @{ N = 'TYPE'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ERROR.TYPE'; F = 'information'; A = @('any'); T = @('#N/A') }

        # date. Kind 'date' is a serial number; 43831 is 2020-01-01.
        @{ N = 'YEAR'; F = 'date'; A = @('date'); T = @('43831') }
        @{ N = 'MONTH'; F = 'date'; A = @('date'); T = @('43831') }
        @{ N = 'DAY'; F = 'date'; A = @('date'); T = @('43831') }
        @{ N = 'HOUR'; F = 'date'; A = @('date'); T = @('0.75') }
        @{ N = 'MINUTE'; F = 'date'; A = @('date'); T = @('0.7521') }
        @{ N = 'SECOND'; F = 'date'; A = @('date'); T = @('0.75210') }
        @{ N = 'WEEKDAY'; F = 'date'; A = @('date', 'int'); T = @('43831', '1'); Min = 1 }
        @{ N = 'WEEKNUM'; F = 'date'; A = @('date', 'int'); T = @('43831', '1'); Min = 1 }
        @{ N = 'EDATE'; F = 'date'; A = @('date', 'int'); T = @('43831', '1') }
        @{ N = 'EOMONTH'; F = 'date'; A = @('date', 'int'); T = @('43831', '1') }
        @{ N = 'DAYS360'; F = 'date'; A = @('date', 'date', 'any'); T = @('43831', '44197', 'FALSE'); Min = 2 }
        @{ N = 'DATEVALUE'; F = 'date'; A = @('text'); T = @('"2020-01-15"') }
        @{ N = 'TIMEVALUE'; F = 'date'; A = @('text'); T = @('"12:30:00"') }
        @{ N = 'TIME'; F = 'date'; A = @('int', 'int', 'int'); T = @('12', '30', '0') }
        @{ N = 'NETWORKDAYS'; F = 'date'; A = @('date', 'date'); T = @('43831', '43861') }
        @{ N = 'WORKDAY'; F = 'date'; A = @('date', 'int'); T = @('43831', '10') }
        @{ N = 'DATEDIF'; F = 'date'; A = @('date', 'date', 'text'); T = @('43831', '44197', '"M"') }
        @{ N = 'YEARFRAC'; F = 'date'; A = @('date', 'date', 'int'); T = @('43831', '44197', '0'); Min = 2 }

        # lookup. S is the table every case of the function starts from.
        # Not listed because XLibur does not implement them yet (they return #NAME?): ADDRESS, CHOOSE, LOOKUP.
        @{ N = 'INDEX'; F = 'lookup'; A = @('ref', 'int', 'int'); T = @('A1:B3', '2', '2'); Min = 2; S = 'A1=1;B1="a";A2=2;B2="b";A3=3;B3="c"' }
        @{ N = 'HLOOKUP'; F = 'lookup'; A = @('any', 'ref', 'int', 'any'); T = @('2', 'A1:C2', '2', 'FALSE'); Min = 3; S = 'A1=1;B1=2;C1=3;A2="a";B2="b";C2="c"' }
        @{ N = 'ROW'; F = 'lookup'; A = @('ref'); T = @('B2'); Min = 0 }
        @{ N = 'COLUMN'; F = 'lookup'; A = @('ref'); T = @('B2'); Min = 0 }
        @{ N = 'ROWS'; F = 'lookup'; A = @('ref'); T = @('A1:B3') }
        @{ N = 'COLUMNS'; F = 'lookup'; A = @('ref'); T = @('A1:B3') }

        # statistical
        @{ N = 'COUNT'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '2'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'COUNTA'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '2'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'COUNTBLANK'; F = 'statistical'; A = @('ref'); T = @('A1:A6'); S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'LARGE'; F = 'statistical'; A = @('ref', 'int'); T = @('A1:A5', '2'); S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'SMALL'; F = 'statistical'; A = @('ref', 'int'); T = @('A1:A5', '2'); S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'RANK'; F = 'statistical'; A = @('num', 'ref', 'int'); T = @('3', 'A1:A5', '0'); Min = 2; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'STDEV'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '10'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'STDEVP'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '10'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'VAR'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '10'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'VARP'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '10'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'MODE'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '1'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'AVERAGEA'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '2'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'MAXA'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '2'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'MINA'; F = 'statistical'; A = @('any', 'any'); T = @('A1:A5', '2'); Min = 1; S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'PERCENTILE'; F = 'statistical'; A = @('ref', 'num'); T = @('A1:A5', '0.3'); S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'QUARTILE'; F = 'statistical'; A = @('ref', 'int'); T = @('A1:A5', '1'); S = 'A1=3;A2=1;A3=4;A4=1;A5=5' }
        @{ N = 'CORREL'; F = 'statistical'; A = @('ref', 'ref'); T = @('A1:A5', 'B1:B5'); S = 'A1=3;A2=1;A3=4;A4=1;A5=5;B1=2;B2=4;B3=5;B4=4;B5=5' }
        @{ N = 'SLOPE'; F = 'statistical'; A = @('ref', 'ref'); T = @('B1:B5', 'A1:A5'); S = 'A1=3;A2=1;A3=4;A4=1;A5=5;B1=2;B2=4;B3=5;B4=4;B5=5' }
        @{ N = 'SUMIF'; F = 'statistical'; A = @('ref', 'any', 'ref'); T = @('A1:A5', '">2"', 'B1:B5'); Min = 2; S = 'A1=3;A2=1;A3=4;A4=1;A5=5;B1=2;B2=4;B3=5;B4=4;B5=5' }
        @{ N = 'AVERAGEIF'; F = 'statistical'; A = @('ref', 'any', 'ref'); T = @('A1:A5', '">2"', 'B1:B5'); Min = 2; S = 'A1=3;A2=1;A3=4;A4=1;A5=5;B1=2;B2=4;B3=5;B4=4;B5=5' }

        # financial
        @{ N = 'PMT'; F = 'financial'; A = @('num', 'int', 'num', 'num', 'int'); T = @('0.05', '10', '1000', '0', '0'); Min = 3 }
        @{ N = 'FV'; F = 'financial'; A = @('num', 'int', 'num', 'num', 'int'); T = @('0.05', '10', '-100', '0', '0'); Min = 3 }
        @{ N = 'PV'; F = 'financial'; A = @('num', 'int', 'num', 'num', 'int'); T = @('0.05', '10', '-100', '0', '0'); Min = 3 }
        @{ N = 'NPER'; F = 'financial'; A = @('num', 'num', 'num', 'num', 'int'); T = @('0.05', '-100', '1000', '0', '0'); Min = 3 }
        @{ N = 'RATE'; F = 'financial'; A = @('int', 'num', 'num', 'num'); T = @('10', '-100', '800', '0'); Min = 3 }
        @{ N = 'IPMT'; F = 'financial'; A = @('num', 'int', 'int', 'num'); T = @('0.05', '1', '10', '1000') }
        @{ N = 'PPMT'; F = 'financial'; A = @('num', 'int', 'int', 'num'); T = @('0.05', '1', '10', '1000') }
        @{ N = 'SLN'; F = 'financial'; A = @('num', 'num', 'num'); T = @('1000', '100', '10') }
        @{ N = 'SYD'; F = 'financial'; A = @('num', 'num', 'num', 'int'); T = @('1000', '100', '10', '1') }
        @{ N = 'DB'; F = 'financial'; A = @('num', 'num', 'int', 'int', 'int'); T = @('1000', '100', '10', '1', '12'); Min = 4 }
        @{ N = 'DDB'; F = 'financial'; A = @('num', 'num', 'int', 'int', 'num'); T = @('1000', '100', '10', '1', '2'); Min = 4 }
        @{ N = 'EFFECT'; F = 'financial'; A = @('num', 'int'); T = @('0.05', '4') }
        @{ N = 'NOMINAL'; F = 'financial'; A = @('num', 'int'); T = @('0.05', '4') }
        @{ N = 'NPV'; F = 'financial'; A = @('num', 'any'); T = @('0.1', 'A1:A3'); S = 'A1=-100;A2=50;A3=70' }
        @{ N = 'IRR'; F = 'financial'; A = @('ref', 'num'); T = @('A1:A3', '0.1'); Min = 1; S = 'A1=-100;A2=50;A3=70' }

        # engineering
        @{ N = 'BIN2DEC'; F = 'engineering'; A = @('text'); T = @('"1010"') }
        @{ N = 'DEC2BIN'; F = 'engineering'; A = @('int', 'int'); T = @('10', '8'); Min = 1 }
        @{ N = 'DEC2HEX'; F = 'engineering'; A = @('int', 'int'); T = @('255', '4'); Min = 1 }
        @{ N = 'DEC2OCT'; F = 'engineering'; A = @('int', 'int'); T = @('64', '4'); Min = 1 }
        @{ N = 'HEX2DEC'; F = 'engineering'; A = @('text'); T = @('"FF"') }
        @{ N = 'OCT2DEC'; F = 'engineering'; A = @('text'); T = @('"17"') }
        @{ N = 'HEX2BIN'; F = 'engineering'; A = @('text', 'int'); T = @('"FF"', '10'); Min = 1 }
        @{ N = 'BIN2HEX'; F = 'engineering'; A = @('text', 'int'); T = @('"1010"', '4'); Min = 1 }
        @{ N = 'DELTA'; F = 'engineering'; A = @('num', 'num'); T = @('5', '5'); Min = 1 }
        @{ N = 'GESTEP'; F = 'engineering'; A = @('num', 'num'); T = @('5', '4'); Min = 1 }
        @{ N = 'ERF'; F = 'engineering'; A = @('num', 'num'); T = @('1', '2'); Min = 1 }
        @{ N = 'ERFC'; F = 'engineering'; A = @('num'); T = @('1') }
        @{ N = 'CONVERT'; F = 'engineering'; A = @('num', 'text', 'text'); T = @('1', '"m"', '"ft"') }
        @{ N = 'COMPLEX'; F = 'engineering'; A = @('num', 'num'); T = @('3', '4') }
        @{ N = 'IMABS'; F = 'engineering'; A = @('text'); T = @('"3+4i"') }
    )
}
