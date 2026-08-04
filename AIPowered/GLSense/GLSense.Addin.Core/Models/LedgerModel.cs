// LedgerModel.cs in GLSense.Addin.Core
// Right-sized carve-out from GLSense\Models\AllModels.cs (FinalWorkingCode) for Group B
// (Cube/Ledger selection - GLCubeDetails + its DataRepository/UserConfig slice):
//   - LedgerModel: DataGrid row model for GLCubeDetails' ledger list (SQLite ledger data
//     merged with the cube-refreshed-date API response).
//   - OperationResult: generic success/message/exception result used by
//     ProcessCubeSelectionNew/Reload.
//   - UserConfigResponse/Preferences/StringToBoolConverter/UserConfigResetResponse: the
//     user-config/get API response consumed by GLCubeDetails.LoadUserPreferencesForCube.
//   - SegmentModel: row model for DataRepository.GetSegments, used by
//     GLCubeDetails.UpdateRibbonForCube to populate the RibSegS ribbon combo after a
//     cube/ledger is selected (SegmentValueModel and the segment *picker* UI itself stay
//     deferred to Group C - only the plain data holder is needed here).
// No logic changes vs. the original.
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json.Serialization;

namespace GLSense.Addin.Core.Models
{
    public class LedgerModel : INotifyPropertyChanged
    {
        public long LedgerId { get; set; }
        public long CubeId { get; set; }
        public string LedgerName { get; set; }
        public long CoaId { get; set; }
        public string PeriodSetName { get; set; }
        public string CurrencyCode { get; set; }
        public string LastRefreshedDate { get; set; }
        public string ADMRefreshedDate { get; set; }
        public string TimeZone { get; set; }

        private bool _hasWarnings;
        public bool HasWarnings
        {
            get => _hasWarnings;
            set
            {
                _hasWarnings = value;
                OnPropertyChanged(nameof(HasWarnings));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // For multi-select
        public bool IsSelected { get; set; }
    }

    public class OperationResult
    {
        public bool IsSuccess { get; set; }
        public string Message { get; set; }
        public Exception Exception { get; set; }
    }

    public class UserConfigResponse
    {
        [JsonPropertyName("preferences")]
        public Preferences Preferences { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("status")]
        public string Status { get; set; }
    }

    public class Preferences
    {
        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("validateCube")]
        public bool? ValidateCube { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("supressZeroBalDrilldown")]
        public bool? SupressZeroBalDrilldown { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runSubLedgerDrilldownAsJob")]
        public bool? RunSubLedgerDrilldownAsJob { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runBalDrilldownAsJob")]
        public bool? RunBalDrilldownAsJob { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runTotalDrilldownAsJob")]
        public bool? RunTotalDrilldownAsJob { get; set; }

        [JsonConverter(typeof(StringToIntConverter))]
        [JsonPropertyName("recordsPerPage")]
        public int? RecordsPerPage { get; set; }

        [JsonConverter(typeof(StringToIntConverter))]
        [JsonPropertyName("refreshCells")]
        public int? RefreshCells { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runJournalDrilldownAsJob")]
        public bool? RunJournalDrilldownAsJob { get; set; }

        [JsonPropertyName("dataOption")]
        public string DataOption { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("journalRealTimeDataEnabled")]
        public bool? JournalRealTimeDataEnabled { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("includeManualJournal")]
        public bool? IncludeManualJournal { get; set; }

        // Ported from GLSense\Models\AllModels.cs (FinalWorkingCode): backs GLUserConfig's
        // "Overwrite drilldown metadata with locally saved" checkbox (Utilities\UserConfig.cs::
        // OverwriteDrilldownMetadata) - when enabled, Drilldowns\DDDatatoWorksheet.cs's
        // ExtractMetadata uses the CustomXMLPart saved via GLDrilldownCustomization's "Save
        // Locally" button (Common\DrilldownMetadataXmlStore.cs) instead of the server's metadata.
        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("overwriteDrilldownMetadata")]
        public bool? OverwriteDrilldownMetadata { get; set; }
    }

    public class UserConfigResetResponse
    {
        public string msg { get; set; }
        public string status { get; set; }
        public string message { get; set; }
    }

    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    public class StringToBoolConverter : System.Text.Json.Serialization.JsonConverter<bool>
    {
        public override bool Read(ref System.Text.Json.Utf8JsonReader reader, Type typeToConvert, System.Text.Json.JsonSerializerOptions options)
        {
            try
            {
                return reader.TokenType switch
                {
                    System.Text.Json.JsonTokenType.True => true,
                    System.Text.Json.JsonTokenType.False => false,
                    System.Text.Json.JsonTokenType.String => reader.GetString()?.ToLower() switch
                    {
                        "true" => true,
                        "false" => false,
                        "1" => true,
                        "0" => false,
                        "yes" => true,
                        "no" => false,
                        _ => throw new System.Text.Json.JsonException($"Cannot convert string '{reader.GetString()}' to bool")
                    },
                    _ => throw new System.Text.Json.JsonException($"Unexpected token type: {reader.TokenType}")
                };
            }
            catch (System.Text.Json.JsonException ex)
            {
                Infrastructure.ServiceLocator.Logger?.LogException(ex, "StringToBoolConverter.Read: failed to convert JSON token to bool");
                throw;
            }
        }

        public override void Write(System.Text.Json.Utf8JsonWriter writer, bool value, System.Text.Json.JsonSerializerOptions options)
        {
            writer.WriteBooleanValue(value);
        }
    }

    // Handles server responses that send numeric fields as JSON strings (e.g.
    // "recordsPerPage":"1", "refreshCells":"1") instead of JSON numbers. Without this,
    // System.Text.Json throws on a plain int/int? property when the token is a string,
    // which was causing ApiResponseHelper.Parse<UserConfigResponse> to fail outright on
    // an otherwise well-formed /user-config preferences response. (Matches the same fix
    // made in FinalWorkingCode's Models\AllModels.cs.)
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    public class StringToIntConverter : System.Text.Json.Serialization.JsonConverter<int>
    {
        public override int Read(ref System.Text.Json.Utf8JsonReader reader, Type typeToConvert, System.Text.Json.JsonSerializerOptions options)
        {
            try
            {
                switch (reader.TokenType)
                {
                    case System.Text.Json.JsonTokenType.Number:
                        return reader.GetInt32();
                    case System.Text.Json.JsonTokenType.String:
                        var s = reader.GetString();
                        if (int.TryParse(s, out var result))
                            return result;
                        throw new System.Text.Json.JsonException($"Cannot convert string '{s}' to int");
                    default:
                        throw new System.Text.Json.JsonException($"Unexpected token type: {reader.TokenType}");
                }
            }
            catch (System.Text.Json.JsonException ex)
            {
                Infrastructure.ServiceLocator.Logger?.LogException(ex, "StringToIntConverter.Read: failed to convert JSON token to int");
                throw;
            }
        }

        public override void Write(System.Text.Json.Utf8JsonWriter writer, int value, System.Text.Json.JsonSerializerOptions options)
        {
            writer.WriteNumberValue(value);
        }
    }

    public class SegmentModel : INotifyPropertyChanged
    {
        private long _id;
        private long _cubeId;
        private long _ledgerId;
        private long _coaId;
        private string _segmentName;
        private long _segmentValueSetId;
        private string _securityEnabledFlag;
        private string _defaultType;
        private string _defaultValue;
        private int _displaySize;
        private string _segmentDelimiter;
        private string _applicationColumnName;
        // Group H (LOVs/Roller/Account dialogs) addition - port of GLSense\Models\
        // AllModels.cs SegmentModel's Value/Reference/IsVisible/IsTextEnabled/
        // IsRefEditEnabled/SelectedValues fields (FinalWorkingCode). Only exercised by
        // SegmentSelectorViewModel's "Ref"-mode branches (GLSegmentRef, still out of
        // scope - see Views\GLAccountsRef.xaml.cs's header comment) and left unused/
        // default by every other consumer of SegmentModel (GLSegmentValues included,
        // which always runs in "val" mode) - added here so SegmentSelectorViewModel
        // compiles against the one shared SegmentModel class, exactly as the old project
        // did.
        private string _value;
        private string _reference;
        private bool _isVisible;
        private bool _isTextEnabled = true;
        private bool _isRefEditEnabled = true;
        private bool _isUserSelected;
        private ObservableCollection<SegmentSelectionModel> _selectedValues = new ObservableCollection<SegmentSelectionModel>();

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        public long Id { get => _id; set => SetProperty(ref _id, value); }
        public long CubeId { get => _cubeId; set => SetProperty(ref _cubeId, value); }
        public long LedgerId { get => _ledgerId; set => SetProperty(ref _ledgerId, value); }
        public long CoaId { get => _coaId; set => SetProperty(ref _coaId, value); }
        public string SegmentName { get => _segmentName; set => SetProperty(ref _segmentName, value); }
        public long SegmentValueSetId { get => _segmentValueSetId; set => SetProperty(ref _segmentValueSetId, value); }
        public string SecurityEnabledFlag { get => _securityEnabledFlag; set => SetProperty(ref _securityEnabledFlag, value); }
        public string DefaultType { get => _defaultType; set => SetProperty(ref _defaultType, value); }
        public string DefaultValue { get => _defaultValue; set => SetProperty(ref _defaultValue, value); }
        public int DisplaySize { get => _displaySize; set => SetProperty(ref _displaySize, value); }
        public string SegmentDelimiter { get => _segmentDelimiter; set => SetProperty(ref _segmentDelimiter, value); }
        public string ApplicationColumnName { get => _applicationColumnName; set => SetProperty(ref _applicationColumnName, value); }
        public string Value { get => _value; set => SetProperty(ref _value, value); }
        public string Reference { get => _reference; set => SetProperty(ref _reference, value); }
        public bool IsVisible { get => _isVisible; set => SetProperty(ref _isVisible, value); }
        public bool IsTextEnabled { get => _isTextEnabled; set => SetProperty(ref _isTextEnabled, value); }
        public bool IsRefEditEnabled { get => _isRefEditEnabled; set => SetProperty(ref _isRefEditEnabled, value); }
        // GLSegmentManager testing feedback: distinguishes "this Value is still the
        // untouched factory default" from "the user actively picked/typed this Value in
        // this session" so the Segments list subtitle (SegmentSummaryConverter) can say
        // "Selected: X" instead of "Default: X" once the user has touched it. Only ever
        // set true from SegmentSelectorViewModel.UpdateRefWindowState's _selectedRight.Any()
        // branch - reached exclusively from the dual-grid Add/Between/NotBetween/Exclude/
        // Remove button handlers and from directly editing the Value textbox, never from
        // initial default-parsing (InitializeSegment/ParseAndSetSegmentValues, which run
        // before this model's PropertyChanged handler is even subscribed) or from
        // switching segments (SelectedSegment's setter restores SelectedValues directly,
        // without going through UpdateRefWindowState at all).
        public bool IsUserSelected { get => _isUserSelected; set => SetProperty(ref _isUserSelected, value); }
        public ObservableCollection<SegmentSelectionModel> SelectedValues { get => _selectedValues; set => SetProperty(ref _selectedValues, value); }

        public override string ToString()
        {
            return $"{SegmentName} ({ApplicationColumnName})";
        }
    }
}
