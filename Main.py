import BEW_extracting_files as bew
import BEC_extracting_files as bec
import EEEP_extracting_files as eee
import time

def main():
    start_time = time.time()
    print ('BEC')
    bec.main()
    print ('BEW')
    bew.main()
    print ('EEE')
    eee.main()
    print('Done! from ', time.asctime(time.localtime(start_time)), ' to ',time.asctime(time.localtime(time.time())))

if __name__=='__main__':
    main()