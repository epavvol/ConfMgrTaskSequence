using System;
using System.Reflection;

namespace VNetDev.ConfMgr.TaskSequence
{
    /// <summary>
    /// Task sequence progress reporting COM automation object (Microsoft.SMS.TsProgressUI) interface.
    /// Both MDT and SMS task sequence environments are supported.
    /// This is enumerable object provides access to Task Sequence environment variables.
    /// </summary>
    public class TsProgressUi
    {
        #region Com Object Instance

        private const string ProgramId = "Microsoft.SMS.TsProgressUI";
        private Type _objectType;
        private object _instance;
        private object Instance => _instance ?? GetInstance();

        private object GetInstance()
        {
            if ((_objectType = Type.GetTypeFromProgID(ProgramId)) != null)
                _instance = Activator.CreateInstance(_objectType);
            return _instance;
        }

        /// <summary>
        /// Returns <c>true</c> if COM object is accessible and instance has been created,
        /// otherwise <c>false</c>.
        /// </summary>
        public bool IsInstanceExists => Instance != null;

        private void ExecuteMethod(string methodName, object[] arguments = null)
        {
            if (!IsInstanceExists)
                throw new InvalidOperationException("Instance is not available");
            var res = _objectType.InvokeMember(methodName,
                BindingFlags.InvokeMethod, null, _instance,
                arguments ?? Array.Empty<object>());
        }

        #endregion

        #region Main

        /// <summary>
        /// Closes open instances of IProgressUI
        /// </summary>
        public void CloseProgressDialog() => ExecuteMethod(nameof(CloseProgressDialog));

        /// <summary>
        /// The ShowActionProgress method displays custom action progress information
        /// in a dialog box while the custom action is running.
        /// </summary>
        /// <param name="orgName">Organization name</param>
        /// <param name="taskSequenceName">Task sequence name</param>
        /// <param name="customTitle">Sub title</param>
        /// <param name="currentAction">Current action</param>
        /// <param name="step">Steps done</param>
        /// <param name="maxStep">Steps total</param>
        /// <param name="actionExecInfo">Sub-action name</param>
        /// <param name="actionExecStep">Sub-action steps done</param>
        /// <param name="actionExecMaxStep">Sub-action steps total</param>
        public void ShowActionProgress(string orgName = null, string taskSequenceName = null, string customTitle = null,
            string currentAction = null, uint? step = null, uint? maxStep = null, string actionExecInfo = null,
            uint? actionExecStep = null, uint? actionExecMaxStep = null) => ExecuteMethod(nameof(ShowActionProgress),
            new object[]
            {
                orgName, taskSequenceName, customTitle, currentAction, step,
                maxStep, actionExecInfo, actionExecStep, actionExecMaxStep
            });

        /// <summary>
        /// Method displays customizable error information in a dialog box.
        /// </summary>
        /// <param name="orgName">Organization name</param>
        /// <param name="taskSequenceName">Task sequence name</param>
        /// <param name="customTitle">Sub title</param>
        /// <param name="errorMessage">Error message</param>
        /// <param name="errorCode">Error code</param>
        /// <param name="timeOut">Timeout in seconds</param>
        /// <param name="restart">Indicates whether the task sequence will restart the computer when the dialog is closed or the timeout expires.</param>
        public void ShowErrorDialog(string orgName = null, string taskSequenceName = null, string customTitle = null,
            string errorMessage = null, uint errorCode = 1u, uint timeOut = 900u, bool restart = false) =>
            ExecuteMethod(nameof(ShowErrorDialog), new object[]
                {orgName, taskSequenceName, customTitle, errorMessage, errorCode, timeOut, restart ? 1 : 0});

        /// <summary>
        /// Displays customizable dialog box.
        /// </summary>
        /// <param name="text">The text displayed in the message box body.</param>
        /// <param name="caption">The text displayed in the message box windows header.</param>
        public void ShowMessage(string text, string caption = "Message") =>
            ExecuteMethod(nameof(ShowMessage), new object[] {text, caption, 0u});

        /// <summary>
        /// Displays customizable dialog box.
        /// </summary>
        /// <param name="text">The text displayed in the message box body.</param>
        /// <param name="caption">The text displayed in the message box windows header.</param>
        /// <param name="messageType">Message type</param>
        public MessageboxResult ShowMessageEx(string text, string caption = "Message", MessageType messageType = MessageType.Ok)
        {
            var result = 0u;
            ExecuteMethod(nameof(ShowMessageEx), new object[] {text, caption, 0u, result});
            return (MessageboxResult)result;
        }

        /// <summary>
        /// Displays customizable reboot warning dialog box.
        /// </summary>
        /// <param name="orgName">Organization name</param>
        /// <param name="taskSequenceName">Task sequence name</param>
        /// <param name="customTitle">Sub title</param>
        /// <param name="rebootMessage">Reboot message</param>
        /// <param name="rebootTimeout">Timeout in seconds</param>
        public void ShowRebootDialog(string orgName = null, string taskSequenceName = null, string customTitle = null,
            string rebootMessage = null, uint rebootTimeout = 0u) => ExecuteMethod(nameof(ShowRebootDialog),
            new object[] {orgName, taskSequenceName, customTitle, rebootMessage, rebootTimeout});

        /// <summary>
        /// Displays message box to prompt a user to swap media.
        /// </summary>
        /// <param name="taskSequenceName">Task sequence name</param>
        /// <param name="mediaNumber">Media number</param>
        public void ShowSwapMediaDialog(string taskSequenceName, uint mediaNumber) =>
            ExecuteMethod(nameof(ShowSwapMediaDialog), new object[] {taskSequenceName, mediaNumber});

        /// <summary>
        /// Displays custom task sequence progress information in a dialog box.
        /// </summary>
        /// <param name="orgName">Organization name</param>
        /// <param name="taskSequenceName">Task sequence name</param>
        /// <param name="customTitle">Sub title</param>
        /// <param name="currentAction">Current action</param>
        /// <param name="step">Steps done</param>
        /// <param name="maxStep">Steps total</param>
        public void ShowTsProgress(string orgName = null, string taskSequenceName = null, string customTitle = null,
            string currentAction = null, uint? step = null, uint? maxStep = null) => ExecuteMethod(
            nameof(ShowTsProgress), new object[]
                {orgName, taskSequenceName, customTitle, currentAction, step, maxStep});

        #endregion
    }
}