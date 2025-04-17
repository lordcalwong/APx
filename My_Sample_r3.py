import pythonnet
import sys, clr 
import csv
import numpy as np
import os

def print_arrays_to_csv(x, y, filename="data.csv", headers=None):
    """
    Prints multiple 1D horizontal arrays as 2D with additional arrays in columns in a CSV file.
    """
    # Ensure all arrays are NumPy arrays for easier handling
    x = np.array(x)
    y = np.array(y)
    # Repack horizontal rows so each row contains elements from corresponding array in columns
    data = list(zip(x, y))
    data = np.array(data)
    # length = data.shape[0]
    # num_dimensions = data.ndim

    # Write the data to a CSV file
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)  # Write the header row
        writer.writerows(data)  # Write the data rows


# APx API Wrapper for Python
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API.dll") 
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API2.dll") 
# COM Wrapper
# clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API3.dll")

from AudioPrecision.API import *
from System.IO import Directory, Path

# # Open APx500 Application
# APx = APx500_Application(APxOperatingMode.SequenceMode, "-Demo -APx500Flex")
# or better to open an existing project that is already configured
filename = "MySampleProject.approjx"
directory = Directory.GetCurrentDirectory()
fullpath = Path.Combine(directory, filename)
APx = APx500_Application()
APx.OpenProject(fullpath)

# MAIN
APx.Visible = True

# How many channels?
maxInputChCount = APx.Version.MaxAnalogInputChannelCount;
maxOutputChCount = APx.Version.MaxAnalogOutputChannelCount;
print (f"Max Output Channels: {maxOutputChCount}")
print (f"Max Input Channels: {maxInputChCount}")

# use all available channels
APx.SignalPathSetup.AnalogInputChannelCount = maxInputChCount
APx.SignalPathSetup.AnalogOutputChannelCount = maxOutputChCount

# setup Reference load for proper wattage calculation
APx.SignalPathSetup.References.AnalogOutputReferences.Watts.Unit = "OHMS"
APx.SignalPathSetup.References.AnalogOutputReferences.Watts.Value = 4

# How to select existing test to be performed by making it active first
APx.ShowMeasurement(0,"Noise (RMS)")
APx.Noise.Level.Checked = True

# Add a new test 'Acoustic Response' makes it active (no need for ShowMeasurement)
APx.AddMeasurement("Signal Path1", MeasurementType.AcousticResponse)

# Setup Signal Path for Acoustic Response; Delete this section if easier to use existing project settings
# APx.Measurement.SweepType.SetValue(LogChirp);
APx.AcousticResponse.GeneratorWithPilot.Frequencies.Start.Value =20;
APx.AcousticResponse.GeneratorWithPilot.Frequencies.Stop.Value = 20000;
APx.AcousticResponse.GeneratorWithPilot.Levelstracking = True;
APx.AcousticResponse.GeneratorWithPilot.Levels.Sweep.SetValue(OutputChannelIndex.Ch1, "0.000 dBrG");
APx.AcousticResponse.GeneratorWithPilot.Durations.Sweep.Value = 0.250;

#Check a couple results to be included in the active sequence
APx.AcousticResponse.Level.Checked = True
APx.AcousticResponse.ThdRatio.Checked = True
#Add a derived result and configure it
smooth = APx.AcousticResponse.Level.AddDerivedResult(MeasurementResultType.Smooth).Result.AsSmoothResult();
smooth.OctaveSmoothing = OctaveSmoothingType.Octave3
smooth.Name = "Smoothed Response"
smooth.Checked = True

#Add an export data sequence step to automatically export the smoothed data when the measurement is run in a sequence.
exportStep = APx.AcousticResponse.SequenceMeasurement.SequenceSteps.ExportResultDataSteps.Add()
exportStep.ResultName = smooth.Name
exportStep.ExportSpecification = "All Points"
current_dir =  os.getcwd()
print("Set sequence step to export data to ", current_dir)
exportStep.FileName = os.path.join(current_dir , "SmoothedResponse.xlsx")
exportStep.Append = False

# Run the sequence, which will run the Acoustic Response measurement
APx = APx500_Application()
APx.Sequence.Run()

### If Exported is too much, target specific data
# Get Frequency Response Gain XY values by result type
thd_xvalues = APx.Sequence[0]["Acoustic Response"].SequenceResults[MeasurementResultType.ThdRatioVsFrequency].GetXValues(InputChannelIndex.Ch1, VerticalAxis.Left, SourceDataType.Measured, 1)
thd_yvalues = APx.Sequence[0]["Acoustic Response"].SequenceResults[MeasurementResultType.ThdRatioVsFrequency].GetYValues(InputChannelIndex.Ch1, VerticalAxis.Left, SourceDataType.Measured, 1)
# or
# Get Frequency Response RMS Level XY values by result name (string)
level_xvaluesCh2 = APx.Sequence[0]["Acoustic Response"].SequenceResults["Smoothed Response"].GetXValues(InputChannelIndex.Ch2, VerticalAxis.Left, SourceDataType.Measured, 1)
level_yvaluesCh2 = APx.Sequence[0]["Acoustic Response"].SequenceResults["Smoothed Response"].GetYValues(InputChannelIndex.Ch2, VerticalAxis.Left, SourceDataType.Measured, 1)

# Write data to file
headers = ["thd_xvalues", "thd_yvalues"]
print_arrays_to_csv(thd_xvalues, thd_yvalues, "THDRatiovsFreq.csv", headers)
headers = ["level_xvaluesCh2", "level_yvaluesCh2"]
print_arrays_to_csv(level_xvaluesCh2, level_yvaluesCh2, "LevelRMS.csv", headers)


# APx.ShowMeasurement(“Signal Path1”, “Acoustic Response”);
# APx.SignalPathSetup.InputConnector.Type = InputConnectorType.ASIO
# # APx.SignalPathSetup.InputDevice = "APx ASIO Loopback"
# APx.SignalPathSetup.OutputConnector.Type = OutputConnectorType.ASIO
# APx.SignalPathSetup.OutputDevice = "APx ASIO Loopback"
# APx.Sequence.ApplyCheckedState(True)
# APx.Sequence["Signal Path1"].SequenceResults["Frequency Response"].Checked = True
# APx.SignalPathSetup.LowpassFilter.A
# APx.SignalPathSetup.WeightingFilter.A-weighting
# APx.AcousticResponse.Level.ExportSpecification = “1000 points“
# APx. AcousticResponse.Level.ExportSpecification = “All Points“
# APx.FrequencyResponse.Level.YAxis.Unit = "dBV";
# ch1xValues = APx.Sequence["Signal Path1"][“MR"].SequenceResults["RMS Level"].GetXValues(0)
# ch1yValues = APx.Sequence["Signal Path1"][“MR"].SequenceResults["RMS Level"].GetYValues(0)
# meterValues = APx.Sequence["Signal Path1"][“THD+N"].SequenceResults[“THD+N Ratio"].GetMeterValues();
# APx.SignalPathSetup.OutputConnector.Type = OutputConnectorType.AnalogBalanced
# APx.SignalPathSetup.InputConnector.Type = InputConnectorType.AnalogBalanced 
# APx.SignalPathSetup.Measure = MeasurandType.Acoustic
# input1 = APx.SignalPathSetup.InputSettings(APxInputSelection.Input1)
# input1.Channels[0].Name = "Mic"
# input1.Channels[0].Sensitivity.Value = 0.011
# APx.MeasurementRecorder.LevelVsTime
# APx.MeasurementRecorder.Graphs["RMS Level (PDM 16)"].Result.AsXYGraph().GetAllXValues
# APx.FrequencyResponse.Level.YAxis.Unit = "dBV";
# var xVals = APx.AcousticResponse.Level.GetAllXValues(InputChannelIndex.Ch1);
# var yVals = APx.AcousticResponse.Level.GetAllYValues(InputChannelIndex.Ch1);
# APx.AcousticResponse.Level.ExportData("c:\\data\\level.xlsx", “All Points");
# APx.Sequence["Signal Path1"][“MR"].SequenceResults[“RMS Level”].Checked = true;
# ch1xValues = APx.Sequence["Signal Path1"][“MR"].SequenceResults["RMS Level"].GetXValues(0)
# ch1yValues = APx.Sequence["Signal Path1"][“MR"].SequenceResults["RMS Level"].GetYValues(0)
# meterValues = APx.Sequence["Signal Path1"][“THD+N"].SequenceResults[“THD+N Ratio"].GetMeterValues();
# SettlingMeterType[] meterTypes = { SettlingMeterType.RmsLevel, SettlingMeterType.ThdNRatio };
# var settledResults = APx.BenchMode.GetSettledMeterReadings(meterTypes);
# var lvlValues = settledResults[SettlingMeterType.RmsLevel].GetValues("Vrms");
# var thdValues = settledResults[SettlingMeterType.ThdNRatio].GetValues("dB");

# APx.BenchMode.Measurements.Recorder.Show()
# APx.BenchMode.Measurements.Recorder.Start()
# APx.BenchMode.Measurements.Recorder.Start();
# while (APx.BenchMode.Measurements.Recorder.IsStarted)
#   Thread.Sleep(100);
# APx.BenchMode.get
# APx.BenchMode.Measurements.Recorder.ExportData();

# bool errorOccurred = APx.BenchMode.Measurements.Recorder.HasError;
# string errorMsg = APx.BenchMode.Measurements.Recorder.LastErrorMessage;
# APError errorCode = APx.BenchMode.Measurements.Recorder.LastErrorCode;


