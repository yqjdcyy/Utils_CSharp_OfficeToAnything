using System;

namespace OfficeToAnything.Error
{
    public class OperationException : Exception
    {
        public OperationException(string message) : base(message)
        {
        }
        public override string ToString()
        {
            return String.Format("无法操作文档：{0}", this.Message);
        }
    }
}
