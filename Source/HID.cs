using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Interop;
using Microsoft.Win32.SafeHandles;
using Win32Bindings.HID;
using Win32Bindings.HID.Defines;
using Win32Bindings.HID.Structs;
using Win32Bindings.kernel32;
using Win32Bindings.kernel32.Defines;
using Win32Bindings.user32;
using Win32Bindings.user32.Defines;
using Win32Bindings.user32.Structs;
using Win32Bindings.WinMessage;
using Window = System.Windows.Window;

namespace RawHidReading
{
    public interface IHID
    {
        bool StopListenInput(HidDevice[] devices);
        bool StartListenInput(HidDevice[] devices);
        event HidInputEventHandler HidEvent;
        HidDevice[] ListDevices();
    }

    public class HID : IHID
    {
        private readonly IntPtr _handle;
        private readonly HwndSource _source;

        public HID(IntPtr windowHandle)
        {
            if (!Win32Bindings.user32.Window.IsWindow(windowHandle))
            {
                throw new ArgumentException("windowHandle is invalid");
            }

            _handle = windowHandle;
            _source = HwndSource.FromHwnd(_handle);
            if (_source == null)
            {
                throw new Exception("Can nnot create HwndSource from windowHandle");
            }

            _source.AddHook(WindowsMessageHandler);
            HidEventInternal += x => { HidEvent?.Invoke(x); };
        }

        public HidDevice[] ListDevices()
        {
            uint deviceCount = 0;
            //get device count
            var res = RawInput.GetRawInputDeviceList(null, ref deviceCount, (uint) Marshal.SizeOf(typeof(RAWINPUTDEVICELIST)));
            if (res == unchecked((uint) -1))
            {
                return new HidDevice[0];
            }

            var ridList = new RAWINPUTDEVICELIST[deviceCount];
            //get device list
            res = RawInput.GetRawInputDeviceList(ridList, ref deviceCount, (uint) Marshal.SizeOf(typeof(RAWINPUTDEVICELIST)));
            if (res != deviceCount)
            {
                return new HidDevice[0];
            }

            var hidDevices = ridList.Select(x =>
            {
                try
                {
                    if (CreateHidDevice(x.hDevice, out var hidDevice, out var error))
                    {
                        return hidDevice;
                    }
                }
                catch (Exception ex)
                {
                }

                return default(HidDevice);
            }).Where(x => x != default(HidDevice));

            return hidDevices.ToArray();
        }

        public event HidInputEventHandler HidEvent;

        public bool StartListenInput(HidDevice[] devices)
        {
            var rids = devices.Select(x => new RAWINPUTDEVICE
            {
                usUsagePage = x.UsagePage,
                usUsage = x.Usage,
                dwFlags = RIDEV.INPUTSINK,
                hwndTarget = _handle
            }).ToArray();

            var structSize = (uint) Marshal.SizeOf(typeof(RAWINPUTDEVICE));
            var isRegistered = RawInput.RegisterRawInputDevices(rids, (uint) rids.Length, structSize);
            var z = WinKernel.GetLastError();
            return isRegistered;
        }

        public bool StopListenInput(HidDevice[] devices)
        {
            var rids = devices.Select(x => new RAWINPUTDEVICE
            {
                usUsagePage = x.UsagePage,
                usUsage = x.Usage,
                dwFlags = RIDEV.REMOVE,
                hwndTarget = _handle
            }).ToArray();

            var structSize = (uint) Marshal.SizeOf(typeof(RAWINPUTDEVICE));
            var isUnRegistered = RawInput.RegisterRawInputDevices(rids, (uint) rids.Length, structSize);
            return isUnRegistered;
        }

        public static IntPtr GetWindowHandle(Window window)
        {
            var interopHelper = new WindowInteropHelper(window);
            return interopHelper.Handle;
        }

        private static IntPtr WindowsMessageHandler(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg != (int) WindowMessage.WM_INPUT)
            {
                //We only process WM_INPUT messages
                return IntPtr.Zero;
            }

            uint dwSize = 0;
            var pData = IntPtr.Zero;
            //get input data size in dwSize
            var sizeGetResult = RawInput.GetRawInputData(lParam, RID.INPUT, pData, ref dwSize, RAWINPUTHEADER.SIZE);
            //return value is 0 if pData is null and getSize successful
            if (sizeGetResult != 0 || dwSize == 0)
            {
                return IntPtr.Zero;
            }

            pData = Marshal.AllocHGlobal((int) dwSize);
            var bytesCopied = RawInput.GetRawInputData(lParam, RID.INPUT, pData, ref dwSize, RAWINPUTHEADER.SIZE);
            //if pData!=null and call successful then bytesCopied==number of bytes copied into buffer
            if (bytesCopied <= 0)
            {
                return IntPtr.Zero;
            }

            var rawInput = (RAWINPUT) Marshal.PtrToStructure(pData, typeof(RAWINPUT));
            var inputReport = new byte[rawInput.data.hid.dwSizeHid];
            for (var i = 0; i < rawInput.data.hid.dwCount; i++)
            {
                //Compute the address from which to copy our HID input
                var hidInputOffset = 0;
                unsafe
                {
                    var source = (byte*) pData;
                    source += RAWINPUTHEADER.SIZE + RAWHID.SIZE + rawInput.data.hid.dwSizeHid * i;
                    hidInputOffset = (int) source;
                }

                //Copy HID input into our buffer
                Marshal.Copy(new IntPtr(hidInputOffset), inputReport, 0, (int) rawInput.data.hid.dwSizeHid);
            }

            HidEventInternal?.Invoke(inputReport);
            //handled = true;

            return IntPtr.Zero;
        }

        private static event HidInputEventHandler HidEventInternal;

        private static bool GetDeviceInfo(IntPtr hDevice, ref RID_DEVICE_INFO deviceInfo, out int errorCode)
        {
            var success = true;
            var deviceInfoBuffer = IntPtr.Zero;
            try
            {
                //Get Device Info
                var deviceInfoSize = (uint) Marshal.SizeOf(typeof(RID_DEVICE_INFO));
                deviceInfoBuffer = Marshal.AllocHGlobal((int) deviceInfoSize);

                var res = RawInput.GetRawInputDeviceInfo(hDevice, RIDI.DEVICEINFO, deviceInfoBuffer, ref deviceInfoSize);
                if (res <= 0)
                {
                    errorCode = Marshal.GetLastWin32Error();
                    Debug.WriteLine("WM_INPUT could not read device info: " + errorCode);
                    return false;
                }

                //Cast our buffer
                deviceInfo = (RID_DEVICE_INFO) Marshal.PtrToStructure(deviceInfoBuffer, typeof(RID_DEVICE_INFO));
            }
            catch
            {
                Debug.WriteLine("GetRawInputData failed!");
                success = false;
            }
            finally
            {
                //Always executes, prevents memory leak
                Marshal.FreeHGlobal(deviceInfoBuffer);
            }

            errorCode = 0;
            return success;
        }

        private static bool CreateHidDevice(IntPtr hRawInputDevice, out HidDevice device, out int errorCode)
        {
            var manufacturer = string.Empty;
            var product = string.Empty;
            ushort productId = 0;
            ushort vendorId = 0;
            ushort version = 0;
            //Fetch various information defining the given HID device
            var name = RawInput.GetRawInputDeviceName(hRawInputDevice);

            //Fetch device info
            var iInfo = new RID_DEVICE_INFO();
            if (!GetDeviceInfo(hRawInputDevice, ref iInfo, out var getDeviceErrorCode))
            {
                device = default(HidDevice);
                errorCode = getDeviceErrorCode;
                return false;
            }

            //Open our device from the device name/path
            var unsafeHandle = FileOperations.CreateFile(name,
                (FileAccess) 0, //none
                FILE_SHARE.READ | FILE_SHARE.WRITE,
                IntPtr.Zero,
                FileMode.OPEN_EXISTING,
                (FILE_ATTRIBUTE) 0x40000000, //.FILE_FLAG_OVERLAPPED,
                IntPtr.Zero
            );

            var handle = new SafeFileHandle(unsafeHandle, true);

            //Check if CreateFile worked
            if (handle.IsInvalid)
            {
                throw new Exception("HidDevice: CreateFile failed: " + Marshal.GetLastWin32Error());
            }

            //Get manufacturer string
            var manufacturerString = new StringBuilder(256);
            if (HidApiBindings.HidD_GetManufacturerString(handle, manufacturerString, manufacturerString.Capacity))
            {
                manufacturer = manufacturerString.ToString();
            }

            //Get product string
            var productString = new StringBuilder(256);
            if (HidApiBindings.HidD_GetProductString(handle, productString, productString.Capacity))
            {
                product = productString.ToString();
            }

            //Get attributes
            var attributes = new HIDD_ATTRIBUTES();
            if (HidApiBindings.HidD_GetAttributes(handle, ref attributes))
            {
                vendorId = attributes.VendorID;
                productId = attributes.ProductID;
                version = attributes.VersionNumber;
            }

            handle.Close();
            var friendlyName = GetFriendlyName(iInfo, product, name, productId);
            var usagePage = iInfo.hid.usUsagePage;
            var usage = iInfo.hid.usUsage;
            device = new HidDevice(manufacturer, product, vendorId, productId, version, friendlyName, usagePage, usage);
            errorCode = 0;
            return true;
        }

        private static Type UsageCollectionType(UsagePage aUsagePage)
        {
            switch (aUsagePage)
            {
                case UsagePage.GenericDesktopControls:
                    return typeof(GenericDesktop);

                case UsagePage.Consumer:
                    return typeof(Consumer);

                case UsagePage.WindowsMediaCenterRemoteControl:
                    return typeof(WindowsMediaCenter);

                default:
                    return null;
            }
        }

        private static string GetFriendlyName(RID_DEVICE_INFO info, string product, string name, ushort productId)
        {
            //Work out proper suffix for our device root node.
            //That allows users to see in a glance what kind of device this is.
            var suffix = "";
            Type usageCollectionType = null;
            var friendlyName = string.Empty;
            if (info.dwType == RIM_TYPE.HID)
            {
                //Process usage page
                if (Enum.IsDefined(typeof(UsagePage), info.hid.usUsagePage))
                {
                    //We know this usage page, add its name
                    var usagePage = (UsagePage) info.hid.usUsagePage;
                    suffix += " ( " + usagePage + ", ";
                    usageCollectionType = UsageCollectionType(usagePage);
                }
                else
                {
                    //We don't know this usage page, add its value
                    suffix += " ( 0x" + info.hid.usUsagePage.ToString("X4") + ", ";
                }

                //Process usage collection
                //We don't know this usage page, add its value
                if (usageCollectionType == null || !Enum.IsDefined(usageCollectionType, info.hid.usUsage))
                {
                    //Show Hexa
                    suffix += "0x" + info.hid.usUsage.ToString("X4") + " )";
                }
                else
                {
                    //We know this usage page, add its name
                    suffix += Enum.GetName(usageCollectionType, info.hid.usUsage) + " )";
                }
            }
            else if (info.dwType == RIM_TYPE.KEYBOARD)
            {
                suffix = " - Keyboard";
            }
            else if (info.dwType == RIM_TYPE.MOUSE)
            {
                suffix = " - Mouse";
            }

            if (product != null && product.Length > 1)
            {
                //This device as a proper name, use it
                friendlyName = product + suffix;
            }
            else
            {
                //Extract friendly name from name
                char[] delimiterChars = {'#', '&'};
                var words = name.Split(delimiterChars);
                if (words.Length >= 2)
                {
                    //Use our name sub-string to describe this device
                    friendlyName = words[1] + " - 0x" + productId.ToString("X4") + suffix;
                }
                else
                {
                    //No proper name just use the device ID instead
                    friendlyName = "0x" + productId.ToString("X4") + suffix;
                }
            }

            return friendlyName;
        }
    }

    public delegate void HidInputEventHandler(byte[] data);
}