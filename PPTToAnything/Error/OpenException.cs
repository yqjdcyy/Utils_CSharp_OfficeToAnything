using System;

namespace OfficeToAnything.Error
{
    public class OpenException : Exception
    {
        public OpenException()
        {
        }

        public OpenException(string message) : base(message)
        {
        }
        public override string ToString()
        {
            return String.Format("无法打开文档：{0}", this.Message);
        }
    }
}
