SET NOCOUNT ON;
GO
/* ---- backup ---- */
IF OBJECT_ID('Price_bak_1405') IS NULL
    SELECT * INTO Price_bak_1405 FROM Price;
GO

BEGIN TRANSACTION;

/* ===== قیمت چاپ (لاستر) - برای ما : کدهای p ===== */
UPDATE Price SET price = '35000'   WHERE RTRIM(size) = 'p10*15';
UPDATE Price SET price = '45000'   WHERE RTRIM(size) = 'p13*18';
UPDATE Price SET price = '65000'   WHERE RTRIM(size) = 'p16*21';
UPDATE Price SET price = '130000'  WHERE RTRIM(size) = 'p20*30';
UPDATE Price SET price = '170000'  WHERE RTRIM(size) = 'p24*30';
UPDATE Price SET price = '260000'  WHERE RTRIM(size) = 'p30*40';
UPDATE Price SET price = '500000'  WHERE RTRIM(size) = 'p40*60';
UPDATE Price SET price = '600000'  WHERE RTRIM(size) = 'p50*60';
UPDATE Price SET price = '700000'  WHERE RTRIM(size) = 'p50*70';
UPDATE Price SET price = '1300000' WHERE RTRIM(size) = 'p60*90';
UPDATE Price SET price = '1685000' WHERE RTRIM(size) = 'p70*100';

/* ===== قیمت شاسی ۸ میل - برای ما : کدهای s ===== */
UPDATE Price SET price = '78000'   WHERE RTRIM(size) = 's13*18';
UPDATE Price SET price = '93000'   WHERE RTRIM(size) = 's16*21';
UPDATE Price SET price = '147000'  WHERE RTRIM(size) = 's20*25';
UPDATE Price SET price = '158000'  WHERE RTRIM(size) = 's20*30';
UPDATE Price SET price = '244000'  WHERE RTRIM(size) = 's24*30';
UPDATE Price SET price = '308000'  WHERE RTRIM(size) = 's30*40';
UPDATE Price SET price = '660000'  WHERE RTRIM(size) = 's40*60';
UPDATE Price SET price = '815000'  WHERE RTRIM(size) = 's50*70';
UPDATE Price SET price = '1120000' WHERE RTRIM(size) = 's60*90';

/* ===== سایزهای جدید چاپ ===== */
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'p30*30')  INSERT INTO Price(size,[count],price,takhfif) VALUES('p30*30', 1,'185000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'p30*45')  INSERT INTO Price(size,[count],price,takhfif) VALUES('p30*45', 1,'300000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'p30*60')  INSERT INTO Price(size,[count],price,takhfif) VALUES('p30*60', 1,'400000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'p40*50')  INSERT INTO Price(size,[count],price,takhfif) VALUES('p40*50', 1,'480000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'p50*100') INSERT INTO Price(size,[count],price,takhfif) VALUES('p50*100',1,'1000000','0');

/* ===== سایزهای جدید شاسی ===== */
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 's9*12')   INSERT INTO Price(size,[count],price,takhfif) VALUES('s9*12',  1,'52000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 's10*15')  INSERT INTO Price(size,[count],price,takhfif) VALUES('s10*15', 1,'52000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 's30*45')  INSERT INTO Price(size,[count],price,takhfif) VALUES('s30*45', 1,'323000','0');

/* ===== لمینت - برای ما : کدهای l (جدید) ===== */
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l10*15')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l10*15', 1,'14500','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l13*18')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l13*18', 1,'22500','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l16*21')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l16*21', 1,'32250','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l20*30')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l20*30', 1,'57500','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l24*30')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l24*30', 1,'69000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l30*30')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l30*30', 1,'86500','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l30*40')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l30*40', 1,'115000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l30*45')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l30*45', 1,'129500','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l30*60')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l30*60', 1,'173000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l40*50')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l40*50', 1,'192000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l40*60')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l40*60', 1,'230500','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l50*60')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l50*60', 1,'288000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l50*70')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l50*70', 1,'336000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l50*100') INSERT INTO Price(size,[count],price,takhfif) VALUES('l50*100',1,'480000','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l60*90')  INSERT INTO Price(size,[count],price,takhfif) VALUES('l60*90', 1,'518500','0');
IF NOT EXISTS(SELECT 1 FROM Price WHERE RTRIM(size) = 'l70*100') INSERT INTO Price(size,[count],price,takhfif) VALUES('l70*100',1,'672000','0');

COMMIT;
GO

SELECT RTRIM(size) AS size, price FROM Price WHERE size LIKE 'p%' OR size LIKE 's%' OR size LIKE 'l%' ORDER BY size;
