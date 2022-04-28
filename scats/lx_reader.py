import glob
import pandas as pd
from tqdm import tqdm
import csv
import os


def find_lx_files(regions, parent_folder='Z:', sub_folder_file="Sys\Sys.lx"):
    lx_files = []
    for r in regions:
        lx_files.append(os.path.join(parent_folder, r, sub_folder_file))
    return lx_files


def combine_lists2(list1, list2):
    list_combined2 = []
    for d1 in list1:
        dict_combined2 = d1.copy()
        sub_system = d1.get("SS")
        for d2 in list2:
            if d2.get("SS") == sub_system:
                dict_combined2.update(d2)
                list_combined2.append(dict_combined2)
                break
    return list_combined2


class LxAnalysis:

    def __init__(self, lx_files, output_file=None, return_dataframe=True, input_folder=None):
        self.lx_files = lx_files
        self.input_folder = input_folder
        self.output_file = output_file
        self.return_dataframe = return_dataframe
        self.analysisType = None
        self.RegionList = []
        self.RegionDict = {}
        self.intInfoDict = {"int": None, "SS": None, "LM": None, "PP1Low": None, "PP1High": None,
                            "PP1OffsetPhase": None, "PP1Slave": None, "PP2Low": None, "PP2High": None,
                            "PP2OffsetPhase": None, "PP2Slave": None,
                            "PP3Low": None, "PP3High": None, "PP3OffsetPhase": None, "PP3Slave": None, "PP4Low": None,
                            "PP4High": None, 'PP4OffsetPhase': None, "PP4Slave": None}
        self.intAllList = []
        self.ssInfoDict = {"int": None, "SS": None, "PP1": None, "PP2": None, "PP3": None, "PP4": None}
        self.allPhaseSequenceList = []
        self.phaseSequenceDict = {"int": None, "Plan1": None, "Plan2": None, "Plan3": None, "Plan4": None,
                                  "Plan5": None, "Plan6": None, "Plan7": None, "Plan8": None}
        self.ssAllList = []
        self.ssDict = {"SS": None, "LCL": None, "HCL": None, "LP1Low": None, "LP1High": None, 'LP1OffsetPhase': None,
                       "LP2Low": None, "LP2High": None, 'LP2OffsetPhase': None,
                       "LP3Low": None, "LP3High": None, 'LP3OffsetPhase': None, "LP4Low": None, "LP4High": None,
                       'LP4OffsetPhase': None,
                       "LP1Intersection": None, "LP2Intersection": None, "LP3Intersection": None,
                       "LP4Intersection": None, "SSInts": None}

    def reset_dictionaries(self):  # Dictionary to be added, type=
        if self.analysisType == "phaseSequence":
            self.allPhaseSequenceList.append(self.phaseSequenceDict.copy())
            self.phaseSequenceDict = {"int": None, "Plan1": None, "Plan2": None, "Plan3": None, "Plan4": None,
                                      "Plan5": None, "Plan6": None, "Plan7": None, "Plan8": None}
        if self.analysisType == "INT":
            self.intAllList.append(self.intInfoDict.copy())
            self.intInfoDict = {"int": None, "SS": None, "LM": None, "PP1Low": None, "PP1High": None,
                                "PP1OffsetPhase": None, "PP1Slave": None, "PP2Low": None, "PP2High": None,
                                "PP2OffsetPhase": None, "PP2Slave": None,
                                "PP3Low": None, "PP3High": None, "PP3OffsetPhase": None, "PP3Slave": None,
                                "PP4Low": None, "PP4High": None, 'PP4OffsetPhase': None, "PP4Slave": None}
        if self.analysisType == "SS" and self.ssDict.get("SS") is not None:
            self.ssAllList.append(self.ssDict.copy())
            self.ssDict = {"SS": None, "LCL": None, "HCL": None, "LP1Low": None, "LP1High": None,
                           'LP1OffsetPhase': None, "LP2Low": None, "LP2High": None, 'LP2OffsetPhase': None,
                           "LP3Low": None, "LP3High": None, 'LP3OffsetPhase': None, "LP4Low": None, "LP4High": None,
                           'LP4OffsetPhase': None,
                           "LP1Intersection": None, "LP2Intersection": None, "LP3Intersection": None,
                           "LP4Intersection": None, "SSInts": None}

    def combine_lists(self, list1, list2):
        list_combined = []
        for d1 in list1:
            intersection = d1.get("int")
            d1["Region"] = self.Region
            for d2 in list2:
                if d2.get("int") == intersection:
                    dict_combined = d1.copy()
                    dict2_combined = d2.copy()
                    dict_combined.update(dict2_combined)
                    list_combined.append(dict_combined)
                    break
        return list_combined

    def is_digit(self, c):
        try:
            int(c)
            return True
        except ValueError:
            return False

    def pp_high(self, pp_string):
        string_count = 0
        for c in pp_string:
            if self.is_digit(c) or c == '-':
                string_count += 1
            else:
                break

        return pp_string[:string_count], pp_string[string_count:]

    def pp_analysis(self, int_):
        if "SL" in int_:
            pp_slave = True
            pp = int_.split("SL")
        else:
            pp_slave = False
            pp = int_.split(",")
        pp_low = pp[0][4:]
        if len(pp) == 2:
            pp_high, offset_phase = self.pp_high(pp[-1])
        else:
            pp_high = None
            offset_phase = None
        if pp_slave:
            pp_slave = pp_high
            pp_high = pp_low
        return pp_low, pp_high, offset_phase, pp_slave

    def lp_high(self, lp_string):
        string_count = 0
        string_start = False
        intersection_found = False
        for c in lp_string:
            if (self.is_digit(c) or c == '-') and string_start is False:
                string_count += 1
                string_end = string_count
            elif not self.is_digit(c) and not intersection_found:
                string_start = True
                string_end += 1
            elif self.is_digit(c):
                intersection_found = True
        return lp_string[:string_count], lp_string[string_count:string_end], lp_string[string_end:]

    def lp_analysis(self, int_):
        lp = int_.split(",")
        lp_low = lp[0][4:]
        if len(lp) == 2:
            lp_high, offset_phase, link_intersection = self.lp_high(lp[-1])
        else:
            lp_high = None
            offset_phase = None
            link_intersection = None
        return lp_low, lp_high, offset_phase, link_intersection

    def identify_ss_intersection(self):
        ss_dict = {}
        for r in self.intAllList:
            ss = r.get("SS")
            if ss != 0:
                intersection = r.get('int')
                if ss in ss_dict:
                    ss_ints = ss_dict.get(ss)
                    ss_ints.append(intersection)
                else:
                    ss_dict[ss] = [intersection]
        for r in self.ssAllList:
            ss = r.get("SS")
            r.update({"SSInts": ss_dict.get(ss)})
        return

    def update_lp_info(self, lp_no, low, high, offset_phase, intersection):
        self.ssDict.update({f"LP{lp_no}Low": low})
        self.ssDict.update({f"LP{lp_no}High": high})
        self.ssDict.update({f"LP{lp_no}OffsetPhase": offset_phase})
        self.ssDict.update({f"LP{lp_no}Intersection": intersection})
        return

    def update_pp_info(self, pp_no, low, high, offset_phase, slave):
        self.intInfoDict.update({f"PP{pp_no}Low": low})
        self.intInfoDict.update({f"PP{pp_no}High": high})
        self.intInfoDict.update({f"PP{pp_no}OffsetPhase": offset_phase})
        self.intInfoDict.update({f"PP{pp_no}Slave": slave})
        return

    def external_intersections(self, plan_number, plan_type):
        for r in self.RegionList:
            other_linked_intersections = []
            lp = "LP" + str(plan_number)
            ext_link = r.get(lp + "Intersection")
            while ext_link is not None:
                if ext_link in other_linked_intersections:
                    break
                else:
                    other_linked_intersections.append(ext_link)
                    if ext_link[-1:] == 'X':
                        try:
                            ext_link = self.RegionDict.get(ext_link[:-1]).get(lp + "Intersection")
                        except:
                            print('external region not included for ' + ext_link)
                            ext_link = None
                    else:
                        try:
                            ext_link = self.RegionDict.get(ext_link).get(lp + "Intersection")
                        except:
                            print('external region not included for ' + ext_link)
                            ext_link = None
            r.update({"otherLinkedInts" + str(plan_number) + str(plan_type): other_linked_intersections})
            offset = (r.get("PP" + str(plan_number) + str(plan_type)))
            if offset is not None:
                offset = float(offset)
                if other_linked_intersections is not None:
                    for int_ in other_linked_intersections:
                        if int_[-1:] == 'X':
                            try:
                                if self.RegionDict.get(int_[:-1]).get(lp + str(plan_type)) is not None:
                                    offset = offset + float(self.RegionDict.get(int_[:-1]).get(lp + str(plan_type)))
                            except:
                                print('external region not included for ' + int_)
                        else:
                            try:
                                if self.RegionDict.get(int_).get(lp + str(plan_type)) is not None:
                                    offset = offset + float(self.RegionDict.get(int_).get(lp + str(plan_type)))
                            except:
                                print('external region not included for ' + int_)
                # "LP1Low":None,"LP1High":None
                r.update({"offset" + str(plan_number) + str(plan_type): offset})
                if r.get('LCL') is not None:
                    updated_offset = offset
                    if plan_type == 'Low':
                        cl = float(r.get('LCL'))
                    elif plan_type == 'High':
                        cl = float(r.get('LCL'))
                    while updated_offset < 0:
                        updated_offset = updated_offset + cl
                    while updated_offset > cl:
                        updated_offset = updated_offset - cl
                    r.update({"updatedOffset" + str(plan_number) + str(plan_type): updated_offset})

    def analyse_lx_files_in_folder(self):
        if self.lx_files is None:
            self.lx_files = glob.glob(f"{self.input_folder}/*.lx")
        for f1 in tqdm(self.lx_files, desc='analysing_lx_files'):
            file1 = open(f1)
            self.Region = None
            self.intAllList = []
            self.allPhaseSequenceList = []
            self.ssAllList = []
            for line in file1:
                tokens = line.split("!")
                for int_ in tokens:
                    if int_[:5] == 'NAME=':
                        self.Region = int_[5:]
                    if int_[:2] == "I=":
                        self.reset_dictionaries()  # included if above is not required
                        sequence_plan = None
                        self.analysisType = 'phaseSequence'
                        phase_sequence = ""
                        self.phaseSequenceDict.update({"int": int_[2:]})
                    elif int_[:5] == 'PLAN=' and self.analysisType == 'phaseSequence':
                        sequence_plan = int_[5:]

                    elif self.analysisType == 'phaseSequence' and int_[1:2] == "=" and line[:2] != "I=":
                        phase_sequence = phase_sequence + int_[0]
                    elif line == '\n' and self.analysisType == 'phaseSequence':
                        if sequence_plan in ['1', '2', '3', '4', '5', '6', '7', '8']:
                            self.phaseSequenceDict.update({f"Plan{sequence_plan}": phase_sequence})
                        phase_sequence = ""
                        sequence_plan = None
                    elif int_[:3] == "INT":
                        self.reset_dictionaries()
                        self.analysisType = "INT"
                        self.intInfoDict.update({"int": int_[4:]})
                    elif int_[:3] == "LM=" and self.analysisType == "INT":
                        self.intInfoDict.update({"LM": int_[3:]})
                    elif int_[:4] in ["PP1=", "PP2=", "PP3=", "PP4="] and self.analysisType == "INT":
                        pp_low, pp_high, pp_offset_phase, pp_slave = self.pp_analysis(int_)
                        self.update_pp_info(int_[2], pp_low, pp_high, pp_offset_phase, pp_slave)
                    elif int_[:3] == "S#=":
                        subsystem = int_[3:]
                        self.intInfoDict.update({"SS": subsystem})

                    #####################################################
                    # SubSystemInfo
                    elif int_[:3] == "SS=" and line[:3] == "SS=":
                        self.reset_dictionaries()
                        self.analysisType = "SS"
                        self.ssDict.update({"SS": int_[3:]})
                    elif line[:3] == "SS=" and int_[:4] == "LCL=":
                        self.ssDict.update({"LCL": int_[4:]})
                    elif line[:3] == "SS=" and int_[:4] == "HCL=":
                        self.ssDict.update({"HCL": int_[4:]})

                    elif int_[:4] in ["LP1=", "LP2=", "LP3=", "LP4="] and self.analysisType == "SS":
                        lp_low, lp_high, lp_offset_phase, lp_intersection = self.lp_analysis(int_)
                        self.update_lp_info(int_[2], lp_low, lp_high, lp_offset_phase, lp_intersection)
            self.ssAllList.append(
                {"SS": '0', "LCL": None, "HCL": None, "LP1Low": None, "LP1High": None, 'LP1OffsetPhase': None,
                 "LP2Low": None, "LP2High": None, 'LP2OffsetPhase': None,
                 "LP3Low": None, "LP3High": None, 'LP3OffsetPhase': None, "LP4Low": None, "LP4High": None,
                 'LP4OffsetPhase': None,
                 "LP1Intersection": None, "LP2Intersection": None, "LP3Intersection": None, "LP4Intersection": None,
                 "SSInts": None})
            self.reset_dictionaries()
            self.identify_ss_intersection()
            intersection_list = self.combine_lists(self.intAllList, self.allPhaseSequenceList)
            sub_system_list = combine_lists2(intersection_list, self.ssAllList)
            for r in sub_system_list:
                self.RegionList.append(r)
                self.RegionDict[r.get("int")] = r
        for i in [1, 2, 3, 4]:
            for p in ["Low", "High"]:
                self.external_intersections(i, p)

        if self.output_file is not None:
            f = open(self.output_file, 'w', newline='')
            headers = list(self.RegionList[0].keys())
            with f:
                writer = csv.DictWriter(f, fieldnames=headers, quoting=csv.QUOTE_ALL)
                writer.writeheader()
                for r in tqdm(self.RegionList, desc='creating output file'):
                    writer.writerow(r)
        if self.return_dataframe:
            print('creating dataframe')
            df = pd.DataFrame(self.RegionList)
            print('lx analysis complete')
            return df
        return
