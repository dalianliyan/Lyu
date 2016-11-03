using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace LeeYan.Util.Core
{

    public class Encrypt : IDisposable
    {
        #region Constructor

        public Encrypt()
        {
            //定义访问数据加密标准 (DES) 
            mCSP = new DESCryptoServiceProvider();
        }

        #endregion

        #region Properties

        /// <summary>
        /// 对称算法变量:算法的加密服务提供程序 (CSP) 版本的包装对象,
        /// DESCryptoServiceProvider是SymmetricAlgorithm的派生类
        /// 
        /// </summary>
        private readonly SymmetricAlgorithm mCSP;

        /// <summary>
        /// 密钥，8位
        /// </summary>
        private const string defaultKey = "fGhjsK78";

        /// <summary>
        /// 初始化向量,24位
        /// </summary>
        private const string defaultCiv = "ss8edKe3rTsfhH43c+ddwKUs";

        #endregion

        #region Methods

        public string EncryptString(string value, string key = "", string civ = "")
        {
            if (String.IsNullOrEmpty(civ))
            {
                civ = defaultCiv;
            }
            if (String.IsNullOrEmpty(key))
            {
                key = defaultKey;
            }
            //CreateEncryptor创建(对称数据)加密对象
            //定义基本的加密转换运算
            //用指定的密钥和初始化向量创建对称数据加密标准
            var ct = mCSP.CreateEncryptor(Encoding.Default.GetBytes(key), Encoding.Default.GetBytes(civ));
            var byt = Encoding.UTF8.GetBytes(value); //将Value字符转换为UTF-8编码的字节序列
            var ms = new MemoryStream(); //创建内存流
            //定义将内存流链接到加密转换的流
            var cs = new CryptoStream(ms, ct, CryptoStreamMode.Write); //将内存流链接到加密转换的流
            cs.Write(byt, 0, byt.Length); //写入内存流
            cs.FlushFinalBlock(); //将缓冲区中的数据写入内存流，并清除缓冲区
            cs.Close(); //释放内存流
            return Convert.ToBase64String(ms.ToArray()); //将内存流转写入字节数组并转换为string字符
        }

        /// <summary>
        /// Encrypts a string using the SHA256 algorithm.
        /// </summary>
        /// <param name="plainMessage">
        /// The plain Message.
        /// </param>
        /// <returns>
        /// The hash password.
        /// </returns>
        public static string HashPassword(string plainMessage)
        {
            var data = Encoding.UTF8.GetBytes(plainMessage);
            using (HashAlgorithm sha = new SHA256Managed())
            {
                sha.TransformFinalBlock(data, 0, data.Length);
                return Convert.ToBase64String(sha.Hash);
            }
        }


        /// <summary>
        /// 解密字符串
        /// </summary>
        /// <param name="value">要解密的字符串</param>
        /// <param name="key">密钥，8位</param>
        /// <param name="civ">初始化向量，24位</param>
        /// <returns>解密后字符串</returns>
        public string DecryptString(string value, string key = "", string civ = "")
        {
            if (String.IsNullOrEmpty(civ))
            {
                civ = defaultCiv;
            }
            if (String.IsNullOrEmpty(key))
            {
                key = defaultKey;
            }
            //定义基本的加密转换运算
            //用指定的密钥和初始化向量创建对称数据解密标准
            var ct = mCSP.CreateDecryptor(Encoding.Default.GetBytes(key), Encoding.Default.GetBytes(civ));
            var byt = Convert.FromBase64String(value); //将Value(Base 64)字符转换成字节数组
            //定义内存流
            var ms = new MemoryStream();
            //定义将数据流链接到加密转换的流
            var cs = new CryptoStream(ms, ct, CryptoStreamMode.Write);
            cs.Write(byt, 0, byt.Length);
            cs.FlushFinalBlock();
            cs.Close();
            return Encoding.UTF8.GetString(ms.ToArray()); //将字节数组中的所有字符解码为一个字符串
        }

        /// <summary>
        /// 执行与释放或重置非托管资源相关的应用程序定义的任务。
        /// </summary>
        public void Dispose()
        {
            mCSP.Dispose();
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// MD5加密
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string GetMd5Code(string text)
        {
            var result = Encoding.Default.GetBytes(text);
            MD5 md5 = new MD5CryptoServiceProvider();
            var oBytes = md5.ComputeHash(result);
            return BitConverter.ToString(oBytes).Replace("-", "");
        }

        #endregion
    }
}
