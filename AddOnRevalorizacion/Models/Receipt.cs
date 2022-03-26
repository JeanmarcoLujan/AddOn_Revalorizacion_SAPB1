using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRevalorizacion.Models
{
    class Receipt
    {
        public DateTime DocDate { get; set; }
        public DateTime TaxDate { get; set; }
        public int DocEntry { get; set; }
        public int LineNum { get; set; }
        public string Itemcode { get; set; }
        public string DocCur { get; set; }
        public double Quantity { get; set; }
        public double QuantityReal { get; set; }
        public double TotalLine { get; set; }
        public string AccountCode { get; set; }
        public string BatchNum { get; set; }
        public string WarehouseCode { get; set; }
        public string CostingCode { get; set; }
        public string CostingCode2 { get; set; }
        public string CostingCode3 { get; set; }
        public string CostingCode4 { get; set; }
        public string CostingCode5 { get; set; }
        public int Revalorizacion { get; set; }
        public int Salida { get; set; }
        public int Entrada { get; set; }
    }
}
