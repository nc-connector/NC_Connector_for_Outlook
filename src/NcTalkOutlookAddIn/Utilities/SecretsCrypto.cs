// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace NcTalkOutlookAddIn.Utilities
{
    internal sealed class SecretsEncryptedPayload
    {
        internal SecretsEncryptedPayload(string encrypted, string iv, string key)
        {
            Encrypted = encrypted ?? string.Empty;
            Iv = iv ?? string.Empty;
            Key = key ?? string.Empty;
        }

        internal string Encrypted { get; private set; }

        internal string Iv { get; private set; }

        internal string Key { get; private set; }
    }

    internal static class SecretsCrypto
    {
        private const string BcryptAesAlgorithm = "AES";
        private const string BcryptChainingMode = "ChainingMode";
        private const string BcryptChainModeGcm = "ChainingModeGCM";
        private const int KeyBytes = 32;
        private const int IvBytes = 12;
        private const int TagBytes = 16;
        private const int AuthModeInfoVersion = 1;

        internal static SecretsEncryptedPayload EncryptToSecretsPayload(string plainText)
        {
            if (plainText == null)
            {
                throw new ArgumentNullException("plainText");
            }

            byte[] key = RandomBytes(KeyBytes);
            byte[] iv = RandomBytes(IvBytes);
            byte[] cipherWithTag = EncryptAesGcm(Encoding.UTF8.GetBytes(plainText), key, iv);

            return new SecretsEncryptedPayload(
                Convert.ToBase64String(cipherWithTag),
                Convert.ToBase64String(iv),
                Convert.ToBase64String(key));
        }

        private static byte[] RandomBytes(int length)
        {
            byte[] data = new byte[length];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(data);
            }
            return data;
        }

        private static byte[] EncryptAesGcm(byte[] plain, byte[] key, byte[] iv)
        {
            IntPtr algorithm = IntPtr.Zero;
            IntPtr keyHandle = IntPtr.Zero;
            IntPtr ivPtr = IntPtr.Zero;
            IntPtr tagPtr = IntPtr.Zero;

            try
            {
                ThrowIfBcryptFailed(BCryptOpenAlgorithmProvider(out algorithm, BcryptAesAlgorithm, null, 0));
                ThrowIfBcryptFailed(BCryptSetProperty(
                    algorithm,
                    BcryptChainingMode,
                    Encoding.Unicode.GetBytes(BcryptChainModeGcm + "\0"),
                    Encoding.Unicode.GetByteCount(BcryptChainModeGcm + "\0"),
                    0));
                ThrowIfBcryptFailed(BCryptGenerateSymmetricKey(
                    algorithm,
                    out keyHandle,
                    IntPtr.Zero,
                    0,
                    key,
                    key.Length,
                    0));

                byte[] cipher = new byte[plain.Length];
                byte[] tag = new byte[TagBytes];
                ivPtr = AllocCopy(iv);
                tagPtr = AllocCopy(tag);

                var authInfo = new BcryptAuthenticatedCipherModeInfo
                {
                    cbSize = Marshal.SizeOf(typeof(BcryptAuthenticatedCipherModeInfo)),
                    dwInfoVersion = AuthModeInfoVersion,
                    pbNonce = ivPtr,
                    cbNonce = iv.Length,
                    pbTag = tagPtr,
                    cbTag = tag.Length
                };

                int bytesDone;
                ThrowIfBcryptFailed(BCryptEncrypt(
                    keyHandle,
                    plain,
                    plain.Length,
                    ref authInfo,
                    IntPtr.Zero,
                    0,
                    cipher,
                    cipher.Length,
                    out bytesDone,
                    0));

                if (bytesDone != cipher.Length)
                {
                    throw new CryptographicException("AES-GCM encryption returned an unexpected byte count.");
                }

                Marshal.Copy(tagPtr, tag, 0, tag.Length);
                byte[] output = new byte[cipher.Length + tag.Length];
                Buffer.BlockCopy(cipher, 0, output, 0, cipher.Length);
                Buffer.BlockCopy(tag, 0, output, cipher.Length, tag.Length);
                return output;
            }
            finally
            {
                if (tagPtr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(tagPtr);
                }
                if (ivPtr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(ivPtr);
                }
                if (keyHandle != IntPtr.Zero)
                {
                    BCryptDestroyKey(keyHandle);
                }
                if (algorithm != IntPtr.Zero)
                {
                    BCryptCloseAlgorithmProvider(algorithm, 0);
                }
            }
        }

        private static IntPtr AllocCopy(byte[] data)
        {
            IntPtr ptr = Marshal.AllocHGlobal(data.Length);
            Marshal.Copy(data, 0, ptr, data.Length);
            return ptr;
        }

        private static void ThrowIfBcryptFailed(int status)
        {
            if (status != 0)
            {
                throw new CryptographicException("BCrypt failed with status 0x" + status.ToString("X8") + ".");
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct BcryptAuthenticatedCipherModeInfo
        {
            internal int cbSize;
            internal int dwInfoVersion;
            internal IntPtr pbNonce;
            internal int cbNonce;
            internal IntPtr pbAuthData;
            internal int cbAuthData;
            internal IntPtr pbTag;
            internal int cbTag;
            internal IntPtr pbMacContext;
            internal int cbMacContext;
            internal int cbAAD;
            internal long cbData;
            internal int dwFlags;
        }

        [DllImport("bcrypt.dll", CharSet = CharSet.Unicode)]
        private static extern int BCryptOpenAlgorithmProvider(
            out IntPtr phAlgorithm,
            string pszAlgId,
            string pszImplementation,
            int dwFlags);

        [DllImport("bcrypt.dll", CharSet = CharSet.Unicode)]
        private static extern int BCryptSetProperty(
            IntPtr hObject,
            string pszProperty,
            byte[] pbInput,
            int cbInput,
            int dwFlags);

        [DllImport("bcrypt.dll")]
        private static extern int BCryptGenerateSymmetricKey(
            IntPtr hAlgorithm,
            out IntPtr phKey,
            IntPtr pbKeyObject,
            int cbKeyObject,
            byte[] pbSecret,
            int cbSecret,
            int dwFlags);

        [DllImport("bcrypt.dll")]
        private static extern int BCryptEncrypt(
            IntPtr hKey,
            byte[] pbInput,
            int cbInput,
            ref BcryptAuthenticatedCipherModeInfo pPaddingInfo,
            IntPtr pbIV,
            int cbIV,
            byte[] pbOutput,
            int cbOutput,
            out int pcbResult,
            int dwFlags);

        [DllImport("bcrypt.dll")]
        private static extern int BCryptDestroyKey(IntPtr hKey);

        [DllImport("bcrypt.dll")]
        private static extern int BCryptCloseAlgorithmProvider(IntPtr hAlgorithm, int dwFlags);
    }
}
