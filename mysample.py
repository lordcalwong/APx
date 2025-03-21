import pythonnet
import sys, clr 
# import os

# API Wrapper for Python
# clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API.dll") 
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API2.dll") 
# COM Wrapper
# clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API3.dll")

from AudioPrecision.API import *
from System.IO import Directory, Path

# #### Open APx500 Application
# APx = APx500_Application(APxOperatingMode.SequenceMode, "-Demo -APx500Flex")
# or
filename = "SampleProject.approjx"
directory = Directory.GetCurrentDirectory()
fullpath = Path.Combine(directory, filename)
APx = APx500_Application()
APx.OpenProject(fullpath)

APx.Visible = True

# Does it have enough channels?
maxOutputChCount = APx.Version.MaxAnalogOutputChannelCount;
maxInputChCount = APx.Version.MaxAnalogInputChannelCount;
print (f"Max Output Channels: {maxOutputChCount}")
print (f"Max Input Channels: {maxInputChCount}")



# APx.SignalPathSetup
# APx.SignalPathSetup.OutputConnector
# APx.AcousticResponse.Level.ExportSpecification = “1000 points“
# APx.FrequencyResponse.Level.YAxis.Unit = "dBV";
# ch1xValues = APx.Sequence["Signal Path1"][“MR"].SequenceResults["RMS Level"].GetXValues(0)
# ch1yValues = APx.Sequence["Signal Path1"][“MR"].SequenceResults["RMS Level"].GetYValues(0)
# meterValues = APx.Sequence["Signal Path1"][“THD+N"].SequenceResults[“THD+N Ratio"].GetMeterValues();

# APx.Sequence.Run()
# APx.AddMeasurement("Signal Path1", MeasurementType.AcousticResponse)
# APx.AcousticResponse.GeneratorWithPilot.Frequencies.Start.Value = 100;
# APx.AcousticResponse.GeneratorWithPilot.Frequencies.Stop.Value = 10000;
# APx.AcousticResponse.GeneratorWithPilot.Levels.Sweep.SetValue(OutputChannelIndex.Ch1, "1 Vrms");
# APx.AcousticResponse.GeneratorWithPilot.Durations.Sweep.Value = 1;

# APx = APx500_Application()
# APx.SignalPathSetup.OutputConnector.Type = OutputConnectorType.AnalogBalanced
# APx.SignalPathSetup.InputConnector.Type = InputConnectorType.AnalogBalanced 
# APx.SignalPathSetup.Measure = MeasurandType.Acoustic
# input1 = APx.SignalPathSetup.InputSettings(APxInputSelection.Input1)
# input1.Channels[0].Name = "Mic"
# input1.Channels[0].Sensitivity.Value = 0.011

# (APxOperatingMode.SequenceMode)
# MeasurementRecorder.LevelVsTime
# APx.MeasurementRecorder.Graphs["RMS Level (PDM 16)"].Result.AsXYGraph().GetAllXValues
# .GetAllYValues

# APx. AcousticResponse.Level.ExportSpecification = “All Points“
# APx. AcousticResponse.Level.ExportSpecification = “1000 points“
# APx.FrequencyResponse.Level.YAxis.Unit = "dBV";

# APx.ShowMeasurement(“Signal Path1”, “Acoustic Response”);
# var xVals = APx.AcousticResponse.Level.GetAllXValues(InputChannelIndex.Ch1);
# var yVals = APx.AcousticResponse.Level.GetAllYValues(InputChannelIndex.Ch1);

# APx.AcousticResponse.Level.ExportData("c:\\data\\level.xlsx", “All Points");
# APx.Sequence.Run();


# APx.Sequence["Signal Path1"][“MR"].SequenceResults[“RMS Level”].Checked = true;
# APx.Sequence.Run()

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
# APx.BenchMode.Measurements.Recorder.ExportData();

# bool errorOccurred = APx.BenchMode.Measurements.Recorder.HasError;
# string errorMsg = APx.BenchMode.Measurements.Recorder.LastErrorMessage;
# APError errorCode = APx.BenchMode.Measurements.Recorder.LastErrorCode;
