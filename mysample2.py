import pythonnet
import sys, clr 
import csv
import numpy as np

def print_arrays_to_csv(*arrays, filename="data.csv", headers=None):
    """
    Prints multiple one-dimensional arrays as columns in a CSV file. Defaults to "data.csv"
    insert headers list of names in first row. Default headers are Column1, Column2, etc.
    """
    # Ensure all arrays are NumPy arrays for easier handling
    arrays = [np.array(arr) for arr in arrays]

    if not arrays:
            print("Error: No arrays provided to print.")
            return  # Or raise an exception, depending on desired behavior

    # Check if all arrays have the same length
    first_length = len(arrays[0])
    if not all(len(arr) == first_length for arr in arrays):
        raise ValueError("All arrays must have the same length.")

    # Determine the number of columns
    num_columns = len(arrays)

    # Create default headers if none are provided
    if headers is None:
        headers = [f"Column{i+1}" for i in range(num_columns)]
    elif len(headers) != num_columns:
        raise ValueError(f"Number of headers ({len(headers)}) must match the number of arrays ({num_columns}).")

    # Create a list of rows, where each row contains elements from corresponding
    # positions in each input array.
    data = list(zip(*arrays))

    # Write the data to a CSV file
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)  # Write the header row
        writer.writerows(data)  # Write the data rows

if __name__ == '__main__':
    # APx API Wrapper for Python
    # clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API.dll") 
    clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API2.dll") 
    # COM Wrapper
    # clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 9.0\\API\\AudioPrecision.API3.dll")

    from AudioPrecision.API import *
    from System.IO import Directory, Path

    # # Open APx500 Application
    # APx = APx500_Application(APxOperatingMode.SequenceMode, "-Demo -APx500Flex")
    # or open and existing project
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

    # Set to max channels
    APx.SignalPathSetup.AnalogInputChannelCount = maxInputChCount
    APx.SignalPathSetup.AnalogOutputChannelCount = maxOutputChCount

    # Select existing test to be performed but make it active first
    APx.ShowMeasurement(0,"Noise (RMS)")
    APx.Noise.Level.Checked = True

    # Add a new test 'Frequency Response' and make it active
    APx.AddMeasurement("Signal Path1", MeasurementType.AcousticResponse)
    APx.AcousticResponse.Level.Checked = True
    # APx.AcousticResponse.GeneratorWithPilot.Frequencies.Start.Value =20;
    # APx.AcousticResponse.GeneratorWithPilot.Frequencies.Stop.Value = 20000;
    APx.AcousticResponse.GeneratorWithPilot.Levels.Sweep.SetValue(OutputChannelIndex.Ch1, "0.000 dBrG");
    APx.AcousticResponse.GeneratorWithPilot.Durations.Sweep.Value = 1;
            
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
    exportStep.FileName = "$(MyDocuments)\\SmoothedResponse.xlsx"
    exportStep.Append = False

    # Run the sequence, which will run the Acoustic Response measurement
    APx = APx500_Application()
    APx.Sequence.Run()

    ### Get results acquired from last run
    # Get Frequency Response Gain XY values by result type
    thd_xvalues = APx.Sequence[0]["Acoustic Response"].SequenceResults[MeasurementResultType.ThdRatioVsFrequency].GetXValues(InputChannelIndex.Ch1, VerticalAxis.Left, SourceDataType.Measured, 1)
    thd_yvalues = APx.Sequence[0]["Acoustic Response"].SequenceResults[MeasurementResultType.ThdRatioVsFrequency].GetYValues(InputChannelIndex.Ch1, VerticalAxis.Left, SourceDataType.Measured, 1)

    # Get Frequency Response RMS Level XY values by result name (string)
    level_xvaluesCh1 = APx.Sequence[0]["Acoustic Response"].SequenceResults["Smoothed Response"].GetXValues(InputChannelIndex.Ch1, VerticalAxis.Left, SourceDataType.Measured, 1)
    level_yvaluesCh1 = APx.Sequence[0]["Acoustic Response"].SequenceResults["Smoothed Response"].GetYValues(InputChannelIndex.Ch1, VerticalAxis.Left, SourceDataType.Measured, 1)
    level_xvaluesCh2 = APx.Sequence[0]["Acoustic Response"].SequenceResults["Smoothed Response"].GetXValues(InputChannelIndex.Ch2, VerticalAxis.Left, SourceDataType.Measured, 1)
    level_yvaluesCh2 = APx.Sequence[0]["Acoustic Response"].SequenceResults["Smoothed Response"].GetYValues(InputChannelIndex.Ch2, VerticalAxis.Left, SourceDataType.Measured, 1)

    # Write data to file
    headers = ["level_xvaluesCh1", "level_yvaluesCh1", "level_xvaluesCh2", "level_yvaluesCh2"]
    print_arrays_to_csv(level_xvaluesCh1, level_yvaluesCh1, level_xvaluesCh2, level_yvaluesCh2, "LevelRMS.csv", headers)



    # # Append measured data to data file
    # with open(os.path.join(save_path , filename), "a") as datafile:
    #     datafile.write(f"{counter:4.0f}, {dt.hour:02d}.{dt.minute:02d}.{dt.second:02d}, {Vmax:.3f}, {Vrms:.3f}\n")
    #     datafile.close()

    # APx.Sequence["Signal Path1"][“MR"].SequenceResults[“RMS Level”].Checked = true;

    # APx.SignalPathSetup.InputConnector.Type = InputConnectorType.ASIO
    # # APx.SignalPathSetup.InputDevice = "APx ASIO Loopback"
    # APx.SignalPathSetup.OutputConnector.Type = OutputConnectorType.ASIO
    # APx.SignalPathSetup.OutputDevice = "APx ASIO Loopback"

    # APx.Sequence.ApplyCheckedState(True)
    # APx.Sequence["Signal Path1"].SequenceResults["Frequency Response"].Checked = True

    # # APx.SignalPathSetup.LowpassFilter.A
    # APx.SignalPathSetup.WeightingFilter.A-weighting

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
    # APx.MeasurementRecorder.LevelVsTime
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
    # APx.BenchMode.get
    # APx.BenchMode.Measurements.Recorder.ExportData();

    # bool errorOccurred = APx.BenchMode.Measurements.Recorder.HasError;
    # string errorMsg = APx.BenchMode.Measurements.Recorder.LastErrorMessage;
    # APError errorCode = APx.BenchMode.Measurements.Recorder.LastErrorCode;


