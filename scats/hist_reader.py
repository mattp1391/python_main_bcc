import os
import numpy
import datetime
from operator import itemgetter
from tqdm import tqdm


def convert_hist_file(file_path, output_folder = None):
    file_name = os.path.basename(file_path)
    file_directory = os.path.dirname(file_path)
    if file_path[-4:] == 'hist':
        np_bytes = numpy.fromfile(file_path, dtype="uint8")
        byte = 0
        date_time_att = None
        event_list = {}
        data_file = None
        while byte <= (len(np_bytes) - 24):
            length_trim = np_bytes[byte]
            if length_trim == 12:
                byte += 9
                bits = numpy.unpackbits(np_bytes[byte:byte + 2])
                day, month, year = extract_date(bits)
                byte += 2
                bits = numpy.unpackbits(np_bytes[byte:byte + 2])
                hour, mins, secs = extract_time(bits)
                byte += 2
                date_time_att = datetime.datetime(1970 + year, month, day, hour, mins, secs)
                if data_file is None:
                    data_file = str(date_time_att.date())
                byte += 1
            else:
                byte_end = byte + length_trim + 1
                while byte < byte_end:
                    byte += 1
                    site_id = (int.from_bytes(np_bytes[byte:byte + 2], byteorder='little'))
                    byte += 2
                    if site_id not in event_list.keys():
                        event_list[site_id] = {'alarms': [], 'phases': []}
                    bits = numpy.unpackbits(np_bytes[byte])

                    record = {
                        'time': str(date_time_att.time()),
                        'alarms': bits,
                        'date_time': date_time_att
                    }
                    event_list[site_id]['alarms'].append(record)
                    byte += 3

                    while byte < byte_end:
                        bits = numpy.unpackbits(np_bytes[byte:byte + 2])
                        delta_time = (bits[0] * 1 + bits[8] * 256 + bits[9] * 128 + bits[10] * 64 +
                                      bits[11] * 32 + bits[12] * 16 + bits[13] * 8 + bits[14] * 4 +
                                      bits[15] * 2)

                        date_time_end = date_time_att + datetime.timedelta(seconds=int(delta_time))
                        phase = (bits[7] * 1 + bits[6] * 2 + bits[5] * 4)
                        spare = (bits[3] * 1 + bits[2] * 2 + bits[1] * 4)
                        gapped = bits[4] * 1
                        byte += 2
                        record = {
                            'time': str(date_time_end.time()),
                            'phase': phase,
                            'duration': int(delta_time),
                            'spare': spare,
                            'date_time': date_time_end,
                            'gapped': gapped
                        }
                        event_list[site_id]['phases'].append(record)
                    byte += 1
        if not os.path.exists(os.path.join(file_directory, file_name + '.data')):
            os.makedirs(os.path.join(file_directory, file_name + '.data'))
        mapphase = ['ZZZ', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'] #A starts at 1, zzz is dummy
        for i, j in event_list.items():
            if output_folder is None:
                name_file_phase = os.path.join(file_directory, f"{file_name}.data", f"{i}_phase{data_file}.txt")
                name_file_alarm = os.path.join(file_directory, f"{file_name}.data", f"{i}_alarm{data_file}.txt")
            else:
                name_file_phase = os.path.join(output_folder, f"{i}_phase{data_file}.txt")
                name_file_alarm = os.path.join(output_folder, f"{i}_alarm{data_file}.txt")
            file_out_phase = open(name_file_phase, 'w')
            file_out_alarm = open(name_file_alarm, 'w')
            file_out_phase.write('phase,duration,start_time,end_time,site_id,gapped,date\n')
            file_out_alarm.write('site_id,time,alarms\n')

            new_list_phase = sorted(j['phases'], key=itemgetter('date_time'))
            new_list__alarm = sorted(j['alarms'], key=itemgetter('date_time'))
            time_old = None

            for i1, j1 in enumerate(new_list_phase):

                if 'gapped' not in j1.keys():
                    j1['gapped'] = 0

                if time_old is None:
                    file_out_phase.write('%s,%d,%s,%s,%d,%d,%s\n' % (mapphase[j1['phase']],
                                                               j1['duration'],
                                                               str((j1[
                                                                        'date_time'] - datetime.timedelta(
                                                                   seconds=j1['duration'])).time()),
                                                               str(j1['date_time'].time()), i,
                                                               j1['gapped'],
                                                               str(j1['date_time'].date()))
                                   )
                    time_old = j1['date_time']
                else:

                    durata = (j1['date_time'] - time_old).total_seconds()
                    file_out_phase.write('%s,%d,%s,%s,%d,%d,%s\n' % (mapphase[j1['phase']], durata,
                                                               str(time_old.time()),
                                                               str(j1['date_time'].time()), i,
                                                               j1['gapped'],
                                                               str(j1['date_time'].date()))
                                   )
                    time_old = j1['date_time']
            file_out_phase.close()
            for i1, j1 in enumerate(new_list__alarm):
                file_out_alarm.write('%d,%s,%s\n' % (i, str(j1['date_time']), str(j1['alarms'])))
            file_out_alarm.close()


def convert_hist_files_in_folder(input_folder, output_folder=None):
    for root, dirs, files in os.walk(input_folder):
        for file in tqdm(files, desc=f'analysing folder: {root}'):
            file_path = os.path.join(root, file)
            convert_hist_file(file_path, output_folder)
    return


def extract_time(bits):
    secs = (bits[7] * 1 + bits[6] * 2 + bits[5] * 4 + bits[4] * 8 + bits[3] * 16) * 2
    mins = bits[2] * 1 + bits[1] * 2 + bits[0] * 4 + bits[15] * 8 + bits[14] * 16 + bits[13] * 32
    hour = bits[12] * 1 + bits[11] * 2 + bits[10] * 4 + bits[9] * 8 + bits[8] * 16
    return hour, mins, secs


def extract_date(bits):
    month = bits[2] + bits[1] * 2 + bits[0] * 4 + bits[15] * 8
    day = bits[7] * 1 + bits[6] * 2 + bits[5] * 4 + bits[4] * 8 + bits[3] * 16
    year = bits[14] * 1 + bits[13] * 2 + bits[12] * 4 + bits[11] * 8 + bits[10] * 16 + bits[9] * 32 + bits[8] * 64
    return day, month, year
