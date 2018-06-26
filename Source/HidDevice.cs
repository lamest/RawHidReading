using System;

namespace RawHidReading
{
    [Serializable]
    public class HidDevice
    {
        public HidDevice()
        {
        }

        public HidDevice(string vendor, string product, ushort vendorId, ushort productId, ushort version, string friendlyName, ushort usagePage, ushort usage)
        {
            Vendor = vendor ?? throw new ArgumentNullException(nameof(vendor));
            Product = product ?? throw new ArgumentNullException(nameof(product));
            VendorId = vendorId;
            ProductId = productId;
            Version = version;
            FriendlyName = friendlyName ?? throw new ArgumentNullException(nameof(friendlyName));
            UsagePage = usagePage;
            Usage = usage;
        }

        public string Vendor { get; }
        public string Product { get; }
        public ushort VendorId { get; }
        public ushort ProductId { get; }
        public ushort Version { get; }
        public string FriendlyName { get; }

        public ushort UsagePage { get; }
        public ushort Usage { get; }
    }
}