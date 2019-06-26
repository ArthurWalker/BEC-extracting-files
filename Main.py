import BEW_extracting_files as bew
import BEC_extracting_files as bec
import EEEP_extracting_files as eee
import time

def main(path_value):
    start_time = time.time()
    print ('BEC')
    bec.main(path_value)
    print ('BEW')
    bew.main(path_value)
    print ('EEE')
    eee.main(path_value)
    print('Done! from ', time.asctime(time.localtime(start_time)), ' to ',time.asctime(time.localtime(time.time())))
