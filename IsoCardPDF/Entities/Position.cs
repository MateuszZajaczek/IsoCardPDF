using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IsoCardPDF.Entities
{
    internal class Position
    {
        public ProfileType ProfileType { get; set; }
        public string Dimension { get; set; }
        
        public int Quantity { get; set; }

        public string Warning { get; set; }

        public Position(ProfileType profileType, string dimension, int quantity)
        {
            ProfileType = profileType;
            Dimension = dimension;
            Quantity = quantity;
            Warning = null; // Explicitly set to null for clarity
        }

        public Position (ProfileType profileType, string dimension, int quantity, string warning)
        {
            ProfileType = profileType;
            Dimension = dimension;
            Quantity = quantity;
            Warning = warning;

        }

    }
}
