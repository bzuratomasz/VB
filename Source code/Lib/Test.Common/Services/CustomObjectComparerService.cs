using System.Collections.Generic;
using Test.Common.Model.Internal;

namespace Test.Common.Services
{
    internal class CustomObjectComparerService : IEqualityComparer<WireDefinition>
    {
        public CustomObjectComparerService() 
        {
        }

        public bool Equals(WireDefinition x, WireDefinition y)
        {
            return x.WireColor == y.WireColor && x.WireCrossSection == y.WireCrossSection && x.WireTypeName == y.WireTypeName && x.ConnectionLocation == y.ConnectionLocation;
        }

        public int GetHashCode(WireDefinition x)
        {
            return x.WireColor.GetHashCode() ^ x.WireCrossSection.GetHashCode() ^ x.WireTypeName.GetHashCode() ^ x.ConnectionLocation.GetHashCode();
        }
    }
}