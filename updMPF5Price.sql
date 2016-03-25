UPDATE tblSecuritiesPalmsSource INNER JOIN tblFinalPriceSheet ON tblSecuritiesPalmsSource.CUSIP = tblFinalPriceSheet.Cusip 
SET    tblSecuritiesPalmsSource.AggPrice = [New Price]
WHERE (((tblSecuritiesPalmsSource.H1)="Seasoned"));
