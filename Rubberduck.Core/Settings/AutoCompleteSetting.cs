﻿using System.Xml.Serialization;
using System.Configuration;
using Rubberduck.AutoComplete;

namespace Rubberduck.Settings
{
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class AutoCompleteSetting
    {
        public AutoCompleteSetting() { /* default ctor required for XML serialization */ }

        public AutoCompleteSetting(IAutoComplete autoComplete)
        {
            Key = autoComplete.GetType().Name;
            IsEnabled = autoComplete.IsEnabled;
        }

        public string Key { get; set; }
        public bool IsEnabled { get; set; }

        [XmlIgnore]
        public string Description => Resources.Settings.SettingsUI.ResourceManager.GetString(Key + "Description");

        public override bool Equals(object obj)
        {
            var other = obj as AutoCompleteSetting;
            return other != null && other.Key == Key;
        }

        public override int GetHashCode()
        {
            return Key?.GetHashCode() ?? 0;
        }
    }
}