[ALLEEG, EEG, CURRENTSET, ALLCOM] = eeglab;
% Load eeglab
EEG.etc.eeglabvers = '14.1.1'; % this tracks which version of EEGLAB is being used, you may ignore it
     edfFile = 'path\No 1 for MI.edf';
     EEG = pop_biosig(edfFile);
     %The edf file path goes here __/\__
     EEG = eeg_checkset( EEG );
     [ALLEEG, EEG] = pac_pop_main(ALLEEG, EEG, CURRENTSET);
     a = EEG.pac.mi
     save('saveTest.txt', 'a','-ascii')
    