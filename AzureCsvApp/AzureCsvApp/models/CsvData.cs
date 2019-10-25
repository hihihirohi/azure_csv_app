using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureCsvApp.models
{
    class CsvDataModel
    {
        public class CsvData
        {
            public string SpCompanyName { get; set; }
            public string SpDomain { get; set; }
            public string SpSubDomain { get; set; }
            public string PurchaseCompanyCode { get; set; }
            public string PurchaseCompanyName { get; set; }
            public string DateofAcquisition { get; set; }
            public string StartUsagePeriod { get; set; }
            public string FinishUsagePeriod { get; set; }
            public string ElementName { get; set; }
            public string CategoyName { get; set; }
            public string SubCategoryName { get; set; }
            public string Area { get; set; }
            public string ResouceId { get; set; }
            public string IdmSubsucriptionId { get; set; }
            public string Tag { get; set; }
            public string UnitName { get; set; }
            public string InstanceDataResourceUri { get; set; }
            public string InstanceDataLocation { get; set; }
            public string InstanceDataPartNumber { get; set; }
            public string InStanceDataOrderNumber { get; set; }
            public string Domain { get; set; }
            public string SubscriptionId { get; set; }
            public decimal PurchaseValue { get; set; }
            public decimal UsafeFee { get; set; }
            public decimal PurchasePrice { get; set; }
        }
    }
}
