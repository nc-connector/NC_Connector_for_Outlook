/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Security.Cryptography;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Shared local password generator fallback used when backend password generation
     * is unavailable or policy responses are missing.
     */
    internal static class PasswordGenerator
    {
        private const string Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
        private const int MinLength = 8;

        internal static string GenerateLocalPassword(int minLength)
        {
            int length = Math.Max(MinLength, minLength);
            var chars = new char[length];
            using (var rng = RandomNumberGenerator.Create())
            {
                byte[] data = new byte[4];
                for (int i = 0; i < chars.Length; i++)
                {
                    rng.GetBytes(data);
                    int index = (int)(BitConverter.ToUInt32(data, 0) % Alphabet.Length);
                    chars[i] = Alphabet[index];
                }
            }

            return new string(chars);
        }
    }
}
