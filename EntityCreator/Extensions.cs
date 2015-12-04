using Microsoft.Xrm.Sdk;

namespace EntityCreator
{
    public static class Extensions
    {
        public static string ToErrorString(this OrganizationServiceFault fault)
        {
            var errorString = string.Empty;
            if (fault != null)
            {
                errorString = fault.Timestamp + "\n" +
                              fault.Message + "\n" +
                              fault.ErrorCode + "\n" +
                              fault.ErrorDetails + "\n" +
                              fault.TraceText + "\n";
                if (fault.InnerFault != null)
                {
                    errorString += fault.InnerFault.ToErrorString();
                }
            }

            return errorString;
        }
    }
}