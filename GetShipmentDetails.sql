CREATE PROCEDURE GetShipmentDetails
AS
BEGIN
SELECT DISTINCT LBP.GrossQuantity,LBP.CreatedTime,LBP.FlowRate,LBP.ShipmentCompartmentID,LBP.BaseProductID,BP.Code AS BaseProductCode, S.ID AS ShipmentID, S.Code AS ShipmentCode, S.ExitTime, B.Code AS BayCode FROM LoadingDetailsBP LBP INNER JOIN BaseProduct BP ON LBP.BaseProductID = BP.ID INNER JOIN ShipmentCompartment SC ON LBP.ShipmentCompartmentID = SC.ID INNER JOIN Shipment S ON SC.ShipmentID = S.ID INNER JOIN  Bay B ON LBP.BayID = B.ID
END;