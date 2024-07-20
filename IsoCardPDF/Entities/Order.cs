using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IsoCardPDF.Entities
{
    internal class Order
    {
        public string OrderId { get; set; }

        public string Date {  get; set; }
        public List<Position> Positions { get; set; }

        public Order(string orderId)
        {
            OrderId = orderId;
            Positions = new List<Position>();

        }


        public void AddPosition(Position position)
        {
            Positions.Add(position);
        }

        public void RemovePosition(int orderId)
        {

        }
    }
}
