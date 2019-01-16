using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.IO;

namespace Encoder.Security
{
    public enum EncoderType
    {
        //可逆編碼(對稱金鑰)
        AES,
        DES,
        RC2,
        TripleDES,

        //可逆編碼(非對稱金鑰)
        RSA,

        //不可逆編碼(雜湊值)
        MD5,
        SHA1,
        SHA256,
        SHA384,
        SHA512
    }

    public class Encoder
    {
        public string Key { get; set; }
        public string IV { get; set; }

        /// <summary>
        /// 產生新的KEY
        /// </summary>
        /// <param name="type">編碼器種類</param>
        public void GenerateKey(EncoderType type)
        {
            switch (type)
            {
                //可逆編碼(對稱金鑰)
                case EncoderType.AES:
                    using (AesCryptoServiceProvider csp = new AesCryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        IV = Convert.ToBase64String(csp.IV);
                        csp.GenerateKey();
                        Key = Convert.ToBase64String(csp.Key);
                    }
                    break;
                case EncoderType.DES:
                    using (DESCryptoServiceProvider csp = new DESCryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        IV = Convert.ToBase64String(csp.IV);
                        csp.GenerateKey();
                        Key = Convert.ToBase64String(csp.Key);
                    }
                    break;
                case EncoderType.RC2:
                    using (RC2CryptoServiceProvider csp = new RC2CryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        IV = Convert.ToBase64String(csp.IV);
                        csp.GenerateKey();
                        Key = Convert.ToBase64String(csp.Key);
                    }
                    break;
                case EncoderType.TripleDES:
                    using (TripleDESCryptoServiceProvider csp = new TripleDESCryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        IV = Convert.ToBase64String(csp.IV);
                        csp.GenerateKey();
                        Key = Convert.ToBase64String(csp.Key);
                    }
                    break;
                //可逆編碼(非對稱金鑰)
                case EncoderType.RSA:
                    using (RSACryptoServiceProvider csp = new RSACryptoServiceProvider())
                    {
                        IV = "";
                        Key = csp.ToXmlString(true);
                    }
                    break;
            }
        }

        /// <summary>
        /// 加密
        /// </summary>
        /// <param name="type">編碼器種類</param>  
        /// <param name="encrypt">加密前字串</param>
        /// <returns>加密後字串</returns>
        public string Encrypt(EncoderType type, string encrypt)
        {
            string ret = "";
            byte[] inputByteArray = Encoding.UTF8.GetBytes(encrypt);

            switch (type)
            {
                //可逆編碼(對稱金鑰)
                case EncoderType.AES:
                    using (AesCryptoServiceProvider csp = new AesCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(Key);
                        byte[] rgbIV = Convert.FromBase64String(IV);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncoderType.DES:
                    using (DESCryptoServiceProvider csp = new DESCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(Key);
                        byte[] rgbIV = Convert.FromBase64String(IV);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncoderType.RC2:
                    using (RC2CryptoServiceProvider csp = new RC2CryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(Key);
                        byte[] rgbIV = Convert.FromBase64String(IV);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncoderType.TripleDES:
                    using (TripleDESCryptoServiceProvider csp = new TripleDESCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(Key);
                        byte[] rgbIV = Convert.FromBase64String(IV);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                //可逆編碼(非對稱金鑰)
                case EncoderType.RSA:
                    using (RSACryptoServiceProvider csp = new RSACryptoServiceProvider())
                    {
                        csp.FromXmlString(Key);
                        ret = Convert.ToBase64String(csp.Encrypt(inputByteArray, false));
                    }
                    break;
                //不可逆編碼(雜湊值)
                case EncoderType.MD5:
                    using (MD5CryptoServiceProvider csp = new MD5CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncoderType.SHA1:
                    using (SHA1CryptoServiceProvider csp = new SHA1CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncoderType.SHA256:
                    using (SHA256CryptoServiceProvider csp = new SHA256CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncoderType.SHA384:
                    using (SHA384CryptoServiceProvider csp = new SHA384CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncoderType.SHA512:
                    using (SHA512CryptoServiceProvider csp = new SHA512CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
            }
            return ret;
        }

        /// <summary>
        /// 解密
        /// </summary>
        /// <param name="type">編碼器種類</param>  
        /// <param name="decrypt">解密前字串</param>
        /// <returns>解密後字串</returns>
        public string Decrypt(EncoderType type, string decrypt)
        {
            string ret = "";
            byte[] inputByteArray = Convert.FromBase64String(decrypt);
            
            switch (type)
            {
                //可逆編碼(對稱金鑰)
                case EncoderType.AES:
                    using (AesCryptoServiceProvider csp = new AesCryptoServiceProvider())
                    {
                        // byte[] rgbKey = Convert.FromBase64String(Key);
                        byte[] rgbKey = Encoding.UTF8.GetBytes(Key);
                         byte[] rgbIV = Convert.FromBase64String(IV);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Encoding.UTF8.GetString(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncoderType.DES:
                    using (DESCryptoServiceProvider csp = new DESCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(Key);
                        byte[] rgbIV = Convert.FromBase64String(IV);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Encoding.UTF8.GetString(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncoderType.RC2:
                    using (RC2CryptoServiceProvider csp = new RC2CryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(Key);
                        byte[] rgbIV = Convert.FromBase64String(IV);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Encoding.UTF8.GetString(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncoderType.TripleDES:
                    using (TripleDESCryptoServiceProvider csp = new TripleDESCryptoServiceProvider())
                    {                         
                        Key = "Do not try to get Database Information from this class";

                        byte[] rgbKey = Encoding.UTF8.GetBytes(Key);
                        byte[] rgMsg = Encoding.UTF8.GetBytes(decrypt);
                        ret = Decrypt(decrypt, Key);
                        //ret = Encoding.UTF8.GetString(Decrypt(decrypt, Key));
                        //byte[] rgbIV = Convert.FromBase64String(IV);

                        //using (MemoryStream ms = new MemoryStream())
                        //{
                        //    using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                        //    {
                        //        cs.Write(inputByteArray, 0, inputByteArray.Length);
                        //        cs.FlushFinalBlock();
                        //        ret = Encoding.UTF8.GetString(ms.ToArray());
                        //    }
                        //}
                    }
                    break;
                //可逆編碼(非對稱金鑰)
                case EncoderType.RSA:
                    using (RSACryptoServiceProvider csp = new RSACryptoServiceProvider())
                    {
                        csp.FromXmlString(Key);
                        ret = Encoding.UTF8.GetString(csp.Decrypt(inputByteArray, false));
                    }
                    break;
            }
            return ret;
        }

        public static string Decrypt(string encryptedData, string secretKey)
        {

            TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
            DES.Key = UTF8Encoding.UTF8.GetBytes(secretKey.Substring(0,24)); //ASCIIEncoding.ASCII.GetBytes(secretKey);
            DES.Mode = CipherMode.ECB;
            DES.Padding = System.Security.Cryptography.PaddingMode.PKCS7;
            ICryptoTransform DESDecrypt = DES.CreateDecryptor();
            string result = "";
            try
             {
                 byte[] Buffer = Convert.FromBase64String(encryptedData);
                result = UTF8Encoding.UTF8.GetString(DESDecrypt.TransformFinalBlock(Buffer, 0, Buffer.Length));
                //result = ASCIIEncoding.ASCII.GetString(DESDecrypt.TransformFinalBlock(Buffer, 0, Buffer.Length));

                 //MemoryStream msDecrypt = new MemoryStream(Buffer);
                 //CryptoStream csDecrypt = new CryptoStream(msDecrypt,
                 //       DES.CreateDecryptor(DES.Key, DES.IV),
                 //       CryptoStreamMode.Read);

                 //// Create buffer to hold the decrypted data.
                 //byte[] fromEncrypt = new byte[Buffer.Length];

                 //// Read the decrypted data out of the crypto stream
                 //// and place it into the temporary buffer.
                 //csDecrypt.Read(fromEncrypt, 0, fromEncrypt.Length);
                 //result = System.Text.Encoding.Default.GetString(fromEncrypt);
             }
                         catch (Exception e)
             {
             }
                         return result;

            //TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
            ////byte[] bytKey = Encoding.UTF8.GetBytes(secretKey);
            //byte[] bytKey = Convert.FromBase64String(secretKey);
            //DES.Mode = System.Security.Cryptography.CipherMode.ECB;

            //DES.Key = bytKey;

            //DES.Padding = PaddingMode.PKCS7;

            //ICryptoTransform DESEncrypt = DES.CreateEncryptor();

            //Byte[] Buffer = System.Text.UTF8Encoding.UTF8.GetBytes(encryptedData);
            //Byte[] Buffer = System.Text.ASCIIEncoding.ASCII.GetBytes(encryptedData);

            //return Convert.ToBase64String(DESEncrypt.TransformFinalBlock(Buffer, 0, Buffer.Length));
            //return DESEncrypt.TransformFinalBlock(Buffer, 0, Buffer.Length);
            // return rijndaelManaged.CreateDecryptor().TransformFinalBlock(encryptedData, 0, encryptedData.Length);
        }

        //public RijndaelManaged GetRijndaelManaged(String secretKey)
        //{
            
        //    //var keyBytes = new byte[16];
        //    //var secretKeyBytes = Encoding.UTF8.GetBytes(secretKey);
        //    //Array.Copy(secretKeyBytes, keyBytes, Math.Min(keyBytes.Length, secretKeyBytes.Length));
        //    //return new RijndaelManaged
        //    //{
        //    //    Mode = CipherMode.ECB,
        //    //    Padding = PaddingMode.PKCS7,
        //    //    KeySize = 128,
        //    //    BlockSize = 128,
        //    //    Key = keyBytes,
        //    //    IV = keyBytes
        //    //};
        //}




    }
}
