using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureCsvApp
{
    class CsvDataMap : CsvHelper.Configuration.ClassMap<CsvData>
    {
        public CsvDataMap()
        {
            Map(x => x.SpCompanyName).Index(0);
            Map(x => x.SpDomain).Index(1);
            Map(x => x.SpSubDomain).Index(2);
            Map(x => x.PurchaseCompanyCode).Index(3);
            Map(x => x.PurchaseCompanyName).Index(4);
            Map(x => x.DateofAcquisition).Index(5);
            Map(x => x.StartUsagePeriod).Index(6);
            Map(x => x.FinishUsagePeriod).Index(7);
            Map(x => x.ElementName).Index(8);
            Map(x => x.CategoyName).Index(9);
            Map(x => x.SubCategoryName).Index(10);
            Map(x => x.Area).Index(11);
            Map(x => x.ResouceId).Index(12);
            Map(x => x.IdmSubsucriptionId).Index(13);
            Map(x => x.Tag).Index(14);
            Map(x => x.UnitName).Index(15);
            Map(x => x.InstanceDataResourceUri).Index(16);
            Map(x => x.InstanceDataLocation).Index(17);
            Map(x => x.InstanceDataPartNumber).Index(18);
            Map(x => x.InStanceDataOrderNumber).Index(19);
            Map(x => x.Domain).Index(20);
            Map(x => x.SubscriptionId).Index(21);
            Map(x => x.PurchaseValue).Index(22);
            Map(x => x.UsafeFee).Index(23);
            Map(x => x.PurchasePrice).Index(24);
        }
    }
}
