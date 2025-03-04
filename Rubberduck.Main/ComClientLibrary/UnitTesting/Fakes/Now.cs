﻿using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Now : FakeBase
    {
        public Now()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetPresentDate");

            InjectDelegate(new NowDelegate(NowCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void NowDelegate(IntPtr retVal);

        public void NowCallback(IntPtr retVal)
        {
            OnCallBack(true);
            if (!TrySetReturnValue())
            {
                TrySetReturnValue(true);
            }
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<NowDelegate>(NativeFunctionAddress);
                nativeCall(retVal);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
