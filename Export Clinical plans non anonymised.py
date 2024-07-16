######################################################################################################################################
#
#	SCRIPT CODE: FM
#
#	SCRIPT TITLE: Export plans from Clinical Database in Raystation
#
#   VERSION: 1.0
#
#   ORIGINAL SCRIPT WRITTEN BY: Fatima Mahmood
#
#   DESCRIPTION & VERSION HISTORY: N/a
#
#	v1.0: (FM) <DESCRIPTION>
#
#                   _____________________________________________________________________________
#                           
#                           SCRIPT VALIDATION DATE IN RAYSTATION SHOULD MATCH FILE DATE
#                   _____________________________________________________________________________
#
######################################################################################################################################


from connect import *
import csv, datetime, sys
import tkinter as tk
from tkinter import filedialog

###############################################################################
# Plan export function
#
# For the selected patient, the specified beamset in the loaded, this function will first try to export
# the beamset normally then will export ignoring these warnings, should this fail.
# Regardless of the result, a log will be returned

def PKExport(dicom_title):
    plan = get_current("Plan")    
    case = get_current('Case')
    beamset = plan.BeamSets[0]
    examination = beamset.GetPlanningExamination()
    # Creates a blank array for warning messages for excel output
    warnings = []
	#Check plan has dose
    if beamset.FractionDose.DoseValues == None:
        success = "Aborted export as no dose"
        #Adds success to warnings array
        return success, warnings
    # Attempts the export for the given inputs and creates messages for the excel export file
    try:
        warnings.append(case.ScriptableDicomExport( \
            Connection = {"Title":dicom_title},
            Examinations = [examination.Name],
            RtStructureSetsForExaminations = [examination.Name],
            BeamSets = [beamset.BeamSetIdentifier()],
            PhysicalBeamSetDoseForBeamSets=[beamset.BeamSetIdentifier()],
            IgnorePreConditionWarnings = True))
        success = "Exported"
    except SystemError as error:
        try:
            warnings.append(case.ScriptableDicomExport( \
                Connection = {"Title":dicom_title},
                Examinations = [examination.Name],
                RtStructureSetsForExaminations = [examination.Name],
                BeamSets = [beamset.BeamSetIdentifier()],
                PhysicalBeamSetDoseForBeamSets=[beamset.BeamSetIdentifier()],
                IgnorePreConditionWarnings = True))
            success = "Exported Ignoring warnings"
        except SystemError as error:
            success = "Failed to export"
            warnings.append(str(error))
    return success, warnings

###############################################################################
# Custom SCRIPT PARAMETERS

#Which DICOM export node to use

dicom_title = 'RAYSTATION' 

#If some patients have not been upgraded to new version this allows this to happen
allow_patient_upgrade = True

#Location of patient list to be used for upload via a file explorer
#root = tk.Tk()
#root.withdraw()
#id_filename = filedialog.askopenfilename()

#Hardcoded location of patient list
id_filename = "S:\Cancer Services - Radiation Physics\Auto-planning\MVision\Scripts\List of patients to export\PR1_Batch_Export_Patients.csv"

#Timestamp for filenames
timestamp = datetime.datetime.now().strftime("%d-%m-%y_%H%M%S")
#filename is the base directory path where the script saves its output log files
filename = "S:\Cancer Services - Radiation Physics\Auto-planning\MVision\Scripts\Batch Export"
#results_filename is the full path including timestamp and file name where the script writes its output log file
result_filename = filename + 'results_' + timestamp + '.csv'
require_plan_approval = True

###############################################################################
# Main function
#
# This function reads the list of patients/beamsets from the input file and attempts
# to open and export each one. The results of each export is output to the output file

patient_db = get_current("PatientDB")
use_index_service = True
plan_list = []      

# Open the input file and read in list, create an entry in plan_list array for each row of csv
try:
    with open(id_filename, 'r') as csvfile:
        file_reader = csv.reader(csvfile)        
        for i, row in enumerate(file_reader):
            plan_list.append(row)
except Exception as e:
    print("Could not read input file, please confirm that you have specified "
        "the 'id_filename' parameter (including directory and extension) "
        "and that the file is a csv file accessible to Raystation" 
        " \n Full error details: " + str(e))
    sys.exit()

# Write column headers to output file
# with open(result_filename, 'ab') as csvfile:
    # writer = csv.writer(csvfile,dialect='excel')
    # writer.writerow(["patient ID","Plan name","Result","Export Notifications / Error details"])

# Loop over all plans in plan_list array
for plan_details in plan_list:
    # Skip iteration if insufficient data in row
    if len(plan_details) < 2:
        print("Skipped row due to missing details")
        continue
    elif len(plan_details) > 2:
        plan_details = [plan_details[0],plan_details[1]]

    # Attempt to find the patient in Raystation - Edit this if you want more data other than ID and plan name
    pt_id,plan_name = plan_details
    ray_pt_info = patient_db.QueryPatientInfo(Filter={"PatientID": pt_id.strip()}, UseIndexService = use_index_service)
    messages = []
    
    # If patient not found moves onto finishing script and closing
    if len(ray_pt_info) == 0:
        result = "Patient not found"
    # One patient found (normal)
    elif len(ray_pt_info) == 1:
        # Attempt to load patient
        pat_loaded = False
        try:
            patient = patient_db.LoadPatient(PatientInfo=ray_pt_info[0], AllowPatientUpgrade = allow_patient_upgrade)
            pat_loaded = True
        # If cannot load patient abort and write to output file
        except SystemError as error:
            result = 'Patient load failed- is index service working?'
            messages.append(str(error))
        # If patient correctly loaded
        if pat_loaded:
            # Search patient for plan to identify associated case
            case = None
            for c in patient.Cases:
                for p in c.TreatmentPlans:
                    if plan_name.strip().lower() in p.Name.strip().lower():
                        case = c
                        #plan name from input csv
                        plan_name = p.Name
            # If could not find plan abort
            if case == None:
                result = "Plan not found"
            # Plan found
            else:      
                # Retrieve plan information
                case.SetCurrent()
                case = get_current("Case")
                examination = get_current("Examination")
                exam_name = examination.Name
                patient.Save()
                plan_info = case.QueryPlanInfo(Filter = {'Name': plan_name})
                load_fail = True
                # Attempt to Load plan
                try:
                    case.LoadPlan(PlanInfo = plan_info[0])
                    plan = case.TreatmentPlans[plan_name]
                    #Check for plan approval (if required)
                    if not require_plan_approval or (plan.Review != None and plan.Review.ApprovalStatus == 'Approved'):
                        load_fail = False
                    else:
                        result = "Could not load plan or plan not approved"
                except Exception as error:
                    result = "Could not load plan or plan not approved"
                    messages.append(str(error))
                # If plan was loaded correctly
                if not load_fail:
                    try:
                        result, warnings = PKExport(dicom_title)
                        # Warn if more than one beamset
                        if plan.BeamSets.Count > 1:
                            result = 'More than one beamset, first beamset exported'
                    except SystemError as error:
                        result = 'Could not export plan'
                        warnings = [str(error)]
                    # Document export notifications
                    for w in warnings:                 
                        messages.append(w)						    	       	                	                      
    # If more than one patient found with matching ID                    
    else:
        result = 'Export abandoned: >1 patient found with ID'

    # Write result to file
    # with open(result_filename, 'ab') as csvfile:
        # writer = csv.writer(csvfile,dialect='excel')
        # writer.writerow([str(pt_id),plan_name,result,''.join(messages)])  