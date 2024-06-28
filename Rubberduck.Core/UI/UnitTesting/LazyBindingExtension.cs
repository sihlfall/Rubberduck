using System;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;

namespace Rubberduck.UI.UnitTesting {

    // Source: Here https://stackoverflow.com/a/48202247/1419315
    [MarkupExtensionReturnType(typeof(object))]
    public class LazyBindingExtension : MarkupExtension {
        public LazyBindingExtension()
        { }

        public LazyBindingExtension(PropertyPath path) : this()
        {
            Path = path;
        }

        #region Properties

        public IValueConverter Converter { get; set; }
        [TypeConverter(typeof(CultureInfoIetfLanguageTagConverter))]
        public CultureInfo ConverterCulture { get; set; }
        public object ConverterParamter { get; set; }
        public string ElementName { get; set; }
        [ConstructorArgument("path")]
        public PropertyPath Path { get; set; }
        public RelativeSource RelativeSource { get; set; }
        public object Source { get; set; }
        public UpdateSourceTrigger UpdateSourceTrigger { get; set; }
        public bool ValidatesOnDataErrors { get; set; }
        public bool ValidatesOnExceptions { get; set; }
        public bool ValidatesOnNotifyDataErrors { get; set; }

        private Binding binding;
        private UIElement bindingTarget;
        private DependencyProperty bindingTargetProperty;

        #endregion

        #region Init

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            var valueProvider = serviceProvider.GetService(typeof(IProvideValueTarget)) as IProvideValueTarget;
            if (valueProvider != null)
            {
                bindingTarget = valueProvider.TargetObject as UIElement;

                if (bindingTarget == null)
                {
                    throw new NotSupportedException($"Target '{valueProvider.TargetObject}' is not valid for a LazyBinding. The LazyBinding target must be a UIElement.");
                }

                bindingTargetProperty = valueProvider.TargetProperty as DependencyProperty;

                if (bindingTargetProperty == null)
                {
                    throw new NotSupportedException($"The property '{valueProvider.TargetProperty}' is not valid for a LazyBinding. The LazyBinding target property must be a DependencyProperty.");
                }

                binding = new Binding
                {
                    Path = Path,
                    Converter = Converter,
                    ConverterCulture = ConverterCulture,
                    ConverterParameter = ConverterParamter
                };

                if (ElementName != null)
                {
                    binding.ElementName = ElementName;
                }

                if (RelativeSource != null)
                {
                    binding.RelativeSource = RelativeSource;
                }

                if (Source != null)
                {
                    binding.Source = Source;
                }

                binding.UpdateSourceTrigger = UpdateSourceTrigger;
                binding.ValidatesOnDataErrors = ValidatesOnDataErrors;
                binding.ValidatesOnExceptions = ValidatesOnExceptions;
                binding.ValidatesOnNotifyDataErrors = ValidatesOnNotifyDataErrors;

                return SetBinding();
            }

            return null;
        }

        public object SetBinding()
        {
            bindingTarget.IsVisibleChanged += UiElement_IsVisibleChanged;

            updateBinding();

            return bindingTarget.GetValue(bindingTargetProperty);
        }

        #endregion

        #region Event Handlers

        private void UiElement_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            updateBinding();
        }

        #endregion

        #region Update Binding

        private void updateBinding()
        {
            if (bindingTarget.IsVisible)
            {
                ConsolidateBinding();
            }
            else
            {
                ClearBinding();
            }
        }

        private bool _isBind;

        private void ConsolidateBinding()
        {
            if (_isBind)
            {
                return;
            }

            _isBind = true;

            BindingOperations.SetBinding(bindingTarget, bindingTargetProperty, binding);
        }

        private void ClearBinding()
        {
            if (!_isBind)
            {
                return;
            }

            BindingOperations.ClearBinding(bindingTarget, bindingTargetProperty);

            _isBind = false;
        }

        #endregion
    }
}
