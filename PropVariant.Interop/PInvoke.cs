using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using HRESULT = System.Int32;

namespace PropVarinatInterop
{
    internal unsafe static class PInvoke
    {
        /// <inheritdoc cref="PropVariantCopy(PROPVARIANT*, PROPVARIANT*)"/>
        internal static unsafe void PropVariantCopy(out PROPVARIANT pvarDest, in PROPVARIANT pvarSrc)
        {
            fixed (PROPVARIANT* pvarSrcLocal = &pvarSrc)
            {
                fixed (PROPVARIANT* pvarDestLocal = &pvarDest)
                {
                    HRESULT __result = PInvoke.PropVariantCopy(pvarDestLocal, pvarSrcLocal);
                    if (__result != 0)
                        throw new Win32Exception(__result);
                }
            }
        }

        /// <summary>Creates a copy of a PROPVARIANT structure.</summary>
        /// <param name="pvarDest">
        /// <para>Type: <b><a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a>*</b> Pointer to the destination <a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a> structure that receives the copy.</para>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-propvariantcopy#parameters">Read more on docs.microsoft.com</see>.</para>
        /// </param>
        /// <param name="pvarSrc">
        /// <para>Type: <b>const <a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a>*</b> Pointer to the source <a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a> structure.</para>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-propvariantcopy#parameters">Read more on docs.microsoft.com</see>.</para>
        /// </param>
        /// <returns>
        /// <para>Type: <b>HRESULT</b> If this function succeeds, it returns <b xmlns:loc="http://microsoft.com/wdcml/l10n">S_OK</b>. Otherwise, it returns an <b xmlns:loc="http://microsoft.com/wdcml/l10n">HRESULT</b> error code.</para>
        /// </returns>
        /// <remarks>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-propvariantcopy">Learn more about this API from docs.microsoft.com</see>.</para>
        /// </remarks>
        [DllImport("Ole32", ExactSpelling = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static extern unsafe HRESULT PropVariantCopy(PROPVARIANT* pvarDest, PROPVARIANT* pvarSrc);

        /// <inheritdoc cref="PropVariantClear(PROPVARIANT*)"/>
        internal static unsafe void PropVariantClear(ref PROPVARIANT pvar)
        {
            fixed (PROPVARIANT* pvarLocal = &pvar)
            {
                HRESULT __result = PInvoke.PropVariantClear(pvarLocal);
                if (__result != 0) 
                    throw new Win32Exception(__result);
            }
        }

        /// <summary>Clears a PROPVARIANT structure.</summary>
        /// <param name="pvar">
        /// <para>Type: <b><a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a>*</b> Pointer to the <a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a> structure to clear. When this function successfully returns, the <b>PROPVARIANT</b> is zeroed and the type is set to VT_EMPTY.</para>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-propvariantclear#parameters">Read more on docs.microsoft.com</see>.</para>
        /// </param>
        /// <returns>
        /// <para>Type: <b>HRESULT</b> If this function succeeds, it returns <b xmlns:loc="http://microsoft.com/wdcml/l10n">S_OK</b>. Otherwise, it returns an <b xmlns:loc="http://microsoft.com/wdcml/l10n">HRESULT</b> error code.</para>
        /// </returns>
        /// <remarks>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-propvariantclear">Learn more about this API from docs.microsoft.com</see>.</para>
        /// </remarks>
        [DllImport("Ole32", ExactSpelling = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static extern unsafe HRESULT PropVariantClear(PROPVARIANT* pvar);

        /// <inheritdoc cref="FreePropVariantArray(uint, PROPVARIANT*)"/>
        internal static unsafe void FreePropVariantArray(Span<PROPVARIANT> rgvars)
        {
            fixed (PROPVARIANT* rgvarsLocal = rgvars)
            {
                HRESULT __result = PInvoke.FreePropVariantArray((uint)rgvars.Length, rgvarsLocal);
                if (__result != 0)
                    throw new Win32Exception(__result);
            }
        }

        /// <summary>Frees the memory and references used by an array of PROPVARIANT structures.</summary>
        /// <param name="cVariants">
        /// <para>Type: <b>ULONG</b> The number of elements in the array specified by <i>rgvars</i>.</para>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-freepropvariantarray#parameters">Read more on docs.microsoft.com</see>.</para>
        /// </param>
        /// <param name="rgvars">
        /// <para>Type: <b><a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a>*</b> Array of <a href="https://docs.microsoft.com/windows/desktop/api/propidl/ns-propidl-propvariant">PROPVARIANT</a> structures to free. When this function successfully returns, the <b>PROPVARIANT</b> structures in the array are zeroed and their type is set to VT_EMPTY.</para>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-freepropvariantarray#parameters">Read more on docs.microsoft.com</see>.</para>
        /// </param>
        /// <returns>
        /// <para>Type: <b>HRESULT</b> If this function succeeds, it returns <b xmlns:loc="http://microsoft.com/wdcml/l10n">S_OK</b>. Otherwise, it returns an <b xmlns:loc="http://microsoft.com/wdcml/l10n">HRESULT</b> error code.</para>
        /// </returns>
        /// <remarks>
        /// <para><see href="https://docs.microsoft.com/windows/win32/api//propidl/nf-propidl-freepropvariantarray">Learn more about this API from docs.microsoft.com</see>.</para>
        /// </remarks>
        [DllImport("Ole32", ExactSpelling = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        internal static extern unsafe HRESULT FreePropVariantArray(uint cVariants, PROPVARIANT* rgvars);

    }

}


