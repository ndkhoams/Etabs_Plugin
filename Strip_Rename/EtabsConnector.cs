using System;
using ETABSv1;

namespace Strip_Rename
{
    public static class EtabsConnector
    {
        public static (cOAPI etabsObject, cSapModel sapModel) AttachToRunningEtabs()
        {
            var etabsObj = (cOAPI)System.Runtime.InteropServices.Marshal
                .GetActiveObject("CSI.ETABS.API.ETABSObject");

            if (etabsObj == null)
                throw new Exception("Không attach được ETABS instance. Hãy mở ETABS trước.");

            var sapModel = etabsObj.SapModel;
            if (sapModel == null)
                throw new Exception("SapModel is null.");

            return (etabsObj, sapModel);
        }
    }
}