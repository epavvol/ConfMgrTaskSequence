namespace VNetDev.ConfMgr.TaskSequence
{
    /// <summary>
    /// Custom task sequence message type
    /// </summary>
    public enum MessageType
    {
        /// <summary>
        /// Ok
        /// </summary>
        Ok = 0,

        /// <summary>
        /// Ok / Cancel
        /// </summary>
        OkCancel = 1,

        /// <summary>
        /// Abort / Retry / Ignore
        /// </summary>
        AbortRetryIgnore = 2,

        /// <summary>
        /// Yes / No / Cancel
        /// </summary>
        YesNoCancel = 3,

        /// <summary>
        /// Yes / No
        /// </summary>
        YesNo = 4,

        /// <summary>
        /// Retry / Cancel
        /// </summary>
        RetryCancel = 5,

        /// <summary>
        /// Cancel / Try again / Continue
        /// </summary>
        CancelTryAgainContinue = 6
    }
}