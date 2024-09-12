namespace CommonHelper.Interface
{
    public interface ICusErroMessageHelper
    {
        /// <summary>
        /// Get Custom Error Message
        /// </summary>
        /// <param name="cusErroCode">The custom error code</param>
        /// <returns>The custom error message</returns>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2024/09/09  1.00    Arvin       Create CusErroCodeHelper
        /// </history>
        public string CusErroCodeHelper(string cusErroCode);
    }
}
