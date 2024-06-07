import os, zipfile, re, shutil
from pydub import AudioSegment

if not os.path.exists('wav_files'):
    os.mkdir('wav_files')
if not os.path.exists('temp'):
    os.mkdir('temp')
if not os.path.exists('output'):
    os.mkdir('output')

print('Extracting...')
audio_extensions = ('.aiff', '.au', '.mid', '.midi', '.mp3', '.m4a', '.mp4', '.wav', '.wma')
init_dir = os.getcwd() ; init_files = os.listdir()
power_dir = os.path.join(init_dir, "powerpoints")
temp_dir = os.path.join(init_dir, "temp")
power_files = os.listdir(power_dir)
break_slide = os.path.join(init_dir,'new_slide.wav')

if not os.path.exists('wav_files'):
    os.mkdir('wav_files')

for power in power_files :
    print(f"Processing {power}")
    if not os.path.exists('temp'):
        os.mkdir('temp')
    if not power.endswith('.zip') and not power.endswith('.pptx') or power == 'base_library.zip' : continue

    base = os.path.splitext(power)[0] ; new_zip = base + '.zip' 
    filepath = os.path.join(power_dir, power)
    new_filepath = os.path.join(temp_dir, new_zip)
    shutil.copyfile(filepath , new_filepath)
    filepath = new_filepath
    with zipfile.ZipFile(filepath) as myzip:

        if myzip.testzip() != None : print("Some of your media files are corrupted")

        else :
            for file in myzip.namelist() :
                if file.endswith(audio_extensions) and 'image' not in file:
                    myzip.extract(file, path = temp_dir)
               
    myzip.close()

    try : os.chdir(os.path.join(temp_dir,'ppt','media'))
    except FileNotFoundError : print('No audio files in {}'.format(power)) ; continue
    
    audio_files = os.listdir()
    for file in audio_files : based = os.path.splitext(file)[0] ; os.rename(file , based + '.wav') 

    wav_audio_files = os.listdir()
    sort_audios = sorted(wav_audio_files, key = lambda x : int(re.search('[0-9]+', x).group()))
    
    repeated = AudioSegment.from_file(break_slide)

    combined_sounds = AudioSegment.empty()
    
    for k, audio in enumerate(sort_audios):
        combined_sounds += AudioSegment.from_file(audio) + repeated
        print('Extracted {}/{} audios'.format(k+1, len(sort_audios)))
        
    name_file = "{}.mp3".format(base.replace(" ", "_")) ; combined_sounds.export(name_file, format="mp3")
    os.rename(os.path.join(os.getcwd(),name_file) , os.path.join(init_dir,'output',name_file))
    
    os.chdir(init_dir); shutil.rmtree('temp', ignore_errors=False)

input('Program terminated, press enter to exit')
