import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk  # Import the ttk module for Combobox
from tqdm import tqdm  # Import tqdm for the loading bar

# Place the header_mapping dictionary here
header_mapping = {
    "AdministrativeData.DataProtection.confidentiality": "Confidentiality",
    "AdministrativeData.DataProtection.justification": "Justification",
    "AdministrativeData.DataProtection.legislations": "Legislations",
    "AdministrativeData.Endpoint": "Endpoint",
    "AdministrativeData.StudyResultType": "StudyResultType",
    "AdministrativeData.PurposeFlag": "PurposeFlag",
    "AdministrativeData.StudyPeriodStartDate": "StudyPeriodStartDate",
    "AdministrativeData.StudyPeriodEndDate": "StudyPeriodEndDate",
    "AdministrativeData.StudyPeriod": "StudyPeriod",
    "AdministrativeData.Reliability": "Reliability",
    "AdministrativeData.RationalReliability": "RationalReliability",
    "AdministrativeData.DataWaiving": "DataWaiving",
    "AdministrativeData.DataWaivingJustification": "DataWaivingJustification",
    "AdministrativeData.JustificationForTypeOfInformation": "JustificationForTypeOfInformation",
    "AdministrativeData.AttachedJustification.AttachedJustification": "AttachedJustification",
    "AdministrativeData.AttachedJustification.ReasonPurpose": "AttachedJustificationReasonPurpose",
    "AdministrativeData.CrossReference.ReasonPurpose": "CrossReferenceReasonPurpose",
    "AdministrativeData.CrossReference.RelatedInformation": "CrossReferenceRelatedInformation",
    "AdministrativeData.CrossReference.Remarks": "CrossReferenceRemarks",
    "DataSource.Reference": "DataSourceReference",
    "DataSource.DataAccess": "DataSourceAccess",
    "DataSource.DataProtectionClaimed": "DataProtectionClaimed",
    "MaterialsAndMethods.ProductType": "ProductType",
    "MaterialsAndMethods.Guideline.Qualifier": "GuidelineQualifier",
    "MaterialsAndMethods.Guideline.Guideline": "Guideline",
    "MaterialsAndMethods.Guideline.VersionRemarks": "GuidelineVersionRemarks",
    "MaterialsAndMethods.Guideline.Deviation": "GuidelineDeviation",
    "MaterialsAndMethods.MethodNoGuideline": "MethodNoGuideline",
    "MaterialsAndMethods.GLPComplianceStatement": "GLPComplianceStatement",
    "MaterialsAndMethods.TestMaterials.TestMaterialInformation": "TestMaterialInformation",
    "MaterialsAndMethods.TestMaterials.AdditionalTestMaterialInformation": "AdditionalTestMaterialInformation",
    "MaterialsAndMethods.TestMaterials.SpecificDetailsOnTestMaterialUsedForTheStudy": "SpecificDetailsOnTestMaterialUsedForTheStudy",
    "MaterialsAndMethods.TestMaterials.SpecificDetailsOnTestMaterialUsedForTheStudyConfidential": "SpecificDetailsOnTestMaterialUsedForTheStudyConfidential",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod": "AnalyticalMethod",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.MethodID": "AnalyticalMethodID",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.RelatedInformation": "AnalyticalMethodRelatedInformation",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.DetailsOnAnalyticalMethods": "AnalyticalMethodDetails",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.CombinationsOfSubstanceAndSamplePortion.AnalyteIdentity": "AnalyteIdentity",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.CombinationsOfSubstanceAndSamplePortion.AnalysedSamplePortionID": "AnalysedSamplePortionID",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.CombinationsOfSubstanceAndSamplePortion.AnalysedSamplePortionDescription": "AnalysedSamplePortionDescription",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.CombinationsOfSubstanceAndSamplePortion.Fortification.FortificationLevel": "FortificationLevel",
    "MaterialsAndMethods.AnalyticalMethods.AnalyticalMethod.CombinationsOfSubstanceAndSamplePortion.Fortification.Recovery": "FortificationRecovery",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.TrialIdNo": "TrialIdNo",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.TestSiteType": "TestSiteType",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.GeographicLocation": "GeographicLocation",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.TrialDeviation": "TrialDeviation",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.Year": "Year ",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.Country": "Country",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.GeographicRegion": "GeographicRegion",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.StateProvince": "StateProvince",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.County": "County",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.City": "City",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.GPSCoordinates": "GPSCoordinates",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.TypeOfCrop": "TypeOfCrop",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.TypeOfTrial": "TypeOfTrial",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.CropGroupingPrimary": "CropGroupingPrimary",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.CropGroup": "CropGroup",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.Crop": "Crop",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.CropCode": "CropCode",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.CropVariety": "CropVariety",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.ReplantNo": "ReplantNo",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.DateOfPlanting": "DateOfPlanting",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.DateOfSeeding": "DateOfSeeding",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.DateOfFloweringBeginning": "DateOfFloweringBeginning",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.DateOfFloweringEnd": "DateOfFloweringEnd",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.DateOfHarvestBegin": "DateOfHarvestBegin",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.DateOfHarvestEnd": "DateOfHarvestEnd",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.CropPlantBackInterval": "CropPlantBackInterval",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.CropInformation": "CropInformation",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.SoilCharacterization": "SoilCharacterization",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.GeographicLocationAndSoilCharacteristics.OtherDetailsOnTestCrops": "OtherDetailsOnTestCrops",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.PlotID": "PlotID",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.ControlPlot": "ControlPlot",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.CorrespondingControlPlotID": "CorrespondingControlPlotID",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.PlotDescription": "PlotDescription",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.EnvironmentalConditions": "EnvironmentalConditions",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.DetailsOnTestSite": "DetailsOnTestSite",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.ApplicationNo": "ApplicationNo",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.BareSoil": "BareSoil",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.GrowthStageCode": "GrowthStageCode",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.GrowthStage": "GrowthStage",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.PlantRowDistance": "PlantRowDistance",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.CrownHeightDuringApplication": "CrownHeightDuringApplication",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.LeafWallArea": "LeafWallArea",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.DateOfApplication": "DateOfApplication",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.MethodOfApplication": "MethodOfApplication",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.SeedingRate": "SeedingRate",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.ThousandGrainWeight": "ThousandGrainWeight",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.TestMaterialInformation": "TestMaterialInformation",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.DescriptionOfTestItem": "DescriptionOfTestItem",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.FormulationType": "FormulationType",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.TradeName": "TradeName",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.RelatedSubstanceInfo": "RelatedSubstanceInfo",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.NameOfAI": "NameOfAI",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.NominalAIContent": "NominalAIContent",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AppliedAmountActual": "AppliedAmountActual",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AppliedAmountAIActual": "AppliedAmountAIActual",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AmountAISeedActual": "AmountAISeedActual",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AppliedAmountCumulative": "AppliedAmountCumulative",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AdjuvantAdded": "AdjuvantAdded",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AmountOfWaterUsedInSpray": "AmountOfWaterUsedInSpray",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AmountOfWaterUsedInSpray": "AmountOfWaterUsedInSpray",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.Application.TestItem.ActiveIngredients.AmountOfWaterUsedInSpray": "AmountOfWaterUsedInSpray",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.Application.OtherDetailsOnApplication": "OtherDetailsOnApplication",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.SamplingAndAnalyticalMethodology.DetailsOnSampleCollection": "DetailsOnSampleCollection",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.SamplingAndAnalyticalMethodology.DetailsOnSampleHandlingAndPreparation": "DetailsOnSampleHandlingAndPreparation",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.SamplingAndAnalysisOfSoil.DetailsOnSamplingOfSoil": "DetailsOnSamplingOfSoil",
    "MaterialsAndMethods.StudyUsePattern.TrialInformation.PlotDescription.Plot.SamplingAndAnalysisOfSoil.DetailsOnAnalyticalMethodologyForSoilResidues": "DetailsOnAnalyticalMethodologyForSoilResidues",
    "MaterialsAndMethods.AnyOtherInformationOnMaterialsAndMethodsInclTables.OtherInformation": "OtherInformation",
    "ResultsAndDiscussion.StorageStability": "StorageStability",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.TrialIDNo": "SummaryOfRadioactiveResiduesInCropsTrialIDNo",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.PlotID": "SummaryOfRadioactiveResiduesInCropsPlotID",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.SamplingID": "SummaryOfRadioactiveResiduesInCropsSamplingID",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.SamplingTiming": "SamplingTiming",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.GrowthStageCode": "GrowthStageCode",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.GrowthStage": "GrowthStage",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.DateOfSampling": "DateOfSampling",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.SamplingInformation": "SamplingInformation",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.SampledMaterialCommodity": "SampledMaterialCommodity",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.SampledMaterialCommodityDescription": "SampledMaterialCommodityDescription",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.MethodID": "ResidueLevelsMethodID",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.AnalyteIdentity": "ResidueLevelsAnalyteIdentity",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.AnalysisSampleDescription": "AnalysisSampleDescription",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.ExtractionDate": "ExtractionDate",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.AnalysisDate": "AnalysisDate",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.StorageStabilityFactor": "StorageStabilityFactor",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.UseOfFactor": "UseOfFactor",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.CorrectionByStorageStability": "CorrectionByStorageStability",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.Recovery": "Recovery",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.CorrectionByRecovery": "CorrectionByRecovery",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.ReferencePortion": "ReferencePortion",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.ResidueLevelMeasured": "ResidueLevelMeasured",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.CalculatedAnalyteIdentity": "CalculatedAnalyteIdentity",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.ResidueLevelCalculated": "ResidueLevelCalculated",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.ResidueLevels.ResidueLevelCorrected": "ResidueLevelCorrected",
    "ResultsAndDiscussion.SummaryOfRadioactiveResiduesInCrops.SamplingAndResidues.TotalMean": "TotalMean",
    "ResultsAndDiscussion.AnyOtherInformationOnResultsInclTables.OtherInformation": "ResultsOtherInformation",
    "OverallRemarksAttachments.RemarksOnResults": "RemarksOnResults",
    "OverallRemarksAttachments.AttachedBackgroundMaterial.Type": "AttachedBackgroundMaterialType",
    "OverallRemarksAttachments.AttachedBackgroundMaterial.AttachedDocument": "AttachedDocument",
    "OverallRemarksAttachments.AttachedBackgroundMaterial.AttachedSanitisedDocsForPublication": "AttachedSanitisedDocsForPublication",
    "OverallRemarksAttachments.AttachedBackgroundMaterial.Remarks": "AttachedBackgroundMaterialRemarks",
    "ApplicantSummaryAndConclusion.InterpretationOfResults": "InterpretationOfResults",
    "ApplicantSummaryAndConclusion.Conclusions": "Conclusions",
    "ApplicantSummaryAndConclusion.ExecutiveSummary": "ExecutiveSummary"
}

def get_header_row():
    header_row = header_row_entry.get()
    try:
        header_row = int(header_row)
        header_rows_span = 1  # Modify this to match your data
        return header_row, header_rows_span
    except ValueError:
        return None, None

# Create a tkinter window for file selection
root = tk.Tk()
root.withdraw()  # Hide the main tkinter window

# Ask the user to select the template file
template_file_path = filedialog.askopenfilename(title="Select Template CSV File", filetypes=[("CSV Files", "*.csv")])

if not template_file_path:
    print("No template file selected. Exiting.")
    exit()

# Read the template file and extract important headers (semicolon-separated)
important_headers = []

try:
    with open(template_file_path, 'r') as template_file:
        for line in template_file:
            headers = line.strip().split(';')
            important_headers.extend([header.strip() for header in headers if header.strip()])
except Exception as e:
    print("Error reading the template file:", str(e))
    exit()

# Display the template headers
print("Template Headers:")
for header in important_headers:
    print(header)

# Create a tkinter window for inputting the header row
header_row_window = tk.Toplevel(root)
header_row_window.title("Header Row Input")

header_row_label = tk.Label(header_row_window, text="Enter the row number where headers are located:")
header_row_label.pack()

header_row_entry = tk.Entry(header_row_window)
header_row_entry.pack()

header_row_button = tk.Button(header_row_window, text="Submit", command=header_row_window.quit)
header_row_button.pack()

header_row_window.mainloop()

selected_header_row, header_rows_span = get_header_row()

if selected_header_row is None:
    print("Invalid header row number. Exiting.")
    exit()

# Ask the user to select the data file (CSV or XLSX)
data_file_path = filedialog.askopenfilename(title="Select Data File", filetypes=[("Data Files", "*.csv;*.xlsx")])

if not data_file_path:
    print("No data file selected. Exiting.")
    exit()

# Load the selected sheet with the specified header row(s)
try:
    if selected_header_row is None or header_rows_span is None:
        print("Invalid header row number or span. Exiting.")
        exit()

    # Read the header data from the specified rows
    header_data = []
    for i in range(header_rows_span):
        header_row_data = pd.read_excel(data_file_path, skiprows=selected_header_row - 1 + i, nrows=1, header=None).values.tolist()[0]
        header_data.extend(header_row_data)

    data = pd.read_excel(data_file_path, header=selected_header_row - 1)


    # Display the header data read from the specified rows
    print("Header Data:")
    print(header_data)

    # Initialize a dictionary for matched data
    matched_data = {}

# Function to clean and standardize header names (case-insensitive and spaces removed)
    def clean_header(header):
        if isinstance(header, str):
            # Convert the header to lowercase and remove spaces
            return header.lower().replace(" ", "").strip()
        else:
            return str(header).strip().lower()

    # Clean and standardize the headers in the data file
    data.columns = [clean_header(header) for header in data.columns]

    
    # Add the code snippet to print the actual headers in your data file
    print("Header Data (datafile):")
    print(data.columns.tolist())

    # Iterate through the header mapping dictionary and find matching data for each template header
    for template_header, data_header in header_mapping.items():
        cleaned_template_header = clean_header(template_header)

        # Check for a case-insensitive match without spaces
        matching_data_header = next((header for header in data.columns if clean_header(header) == cleaned_template_header), None)

        if matching_data_header:
            data_values = data[matching_data_header].tolist()
            matched_data[template_header] = data_values
            print(f"Template Header: {template_header}, Data Header: {matching_data_header}, Data Values: {data_values}")
        else:
            print(f"No matching data found for template header: {template_header}")


except Exception as e:
    print("Error reading the data file:", str(e))
    exit()