using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace VNetDev.ConfMgr.TaskSequence
{
    /// <summary>
    /// Task sequence environment COM automation object (Microsoft.SMS.TSEnvironment) interface.
    /// Both MDT and SMS task sequence environments are supported.
    /// This helps to report progress to the task sequencing environment.
    /// </summary>
    public class TsEnvironment : IEnumerable<KeyValuePair<string, string>>
    {
        #region Com Object Instance

        private const string ProgramId = "Microsoft.SMS.TSEnvironment";
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

        #endregion

        #region Main

        /// <summary>
        /// Gets list of available task sequence variables
        /// </summary>
        /// <returns>List of variable names</returns>
        public string[] GetVariableList() => (_objectType?
                .InvokeMember(
                    "GetVariables",
                    BindingFlags.InvokeMethod,
                    null,
                    Instance,
                    null) as object[])?
            .Where(x => x is string)
            .Cast<string>()
            .ToArray() ?? Array.Empty<string>();

        #endregion

        #region Indexer

        /// <summary>
        /// Task sequence variable indexed access.
        /// </summary>
        /// <param name="variableName">Name of task sequence variable</param>
        /// <exception cref="InvalidOperationException">Throws if COM object instance is not available</exception>
        public string this[string variableName]
        {
            get => GetVariableValue(variableName);
            set => SetVariableValue(variableName, value);
        }

        private string GetVariableValue(string varName) => IsInstanceExists
            ? _objectType.InvokeMember(
                "Value",
                BindingFlags.GetProperty,
                null,
                _instance,
                new object[] {varName}) as string
            : throw new InvalidOperationException("Object instance is not available!");

        /*private void SetVariableValue(string varName, string varValue) => IsInstanceExists
            ? _objectType.InvokeMember(
                "Value",
                BindingFlags.SetProperty,
                null,
                _instance,
                new object[0])
            : throw new InvalidOperationException("Object instance is not available!");*/

        private void SetVariableValue(string name, string value)
        {
            if (!IsInstanceExists)
            {
                throw new Exception("");
            }

            _objectType.InvokeMember("Value", BindingFlags.SetProperty, null,
                _instance, new object[] {name, value});
        }

        #endregion

        #region IEnumerable implementation

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerator<KeyValuePair<string, string>> GetEnumerator() => GetVariableList()
            .Select(varName => new KeyValuePair<string, string>(varName, this[varName]))
            .GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        #endregion

        #region Casts

        /// <summary>
        /// Implicitly converts TsEnvironment to <c>Dictionary[string, string]</c> object
        /// </summary>
        /// <param name="tsEnvironment">Source TsEnvironment object</param>
        /// <returns></returns>
        public static implicit operator Dictionary<string, string>(TsEnvironment tsEnvironment) =>
            tsEnvironment.ToDictionary(x => x.Key,
                x => x.Value);

        #endregion
    }
}