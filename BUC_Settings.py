import subprocess   #
import shlex        # split commands for subprocess
import time
import argparse		# Get arguments from the command line
import os
import sys
import psutil
from pathlib import Path	# handles path
import pandas as pd
import datetime
from openpyxl import Workbook
import pathlib
import csv
import openpyxl
import pyautogui
import threading
import json




PY_COMMAND = None # Override for py command if the automated command doesn't work
TESTING = False # True stops the pnp script from running, the HMD being restarted, and some spammy commands
VERBOSE = False # will be overwritten by the -v argument



# Set the path to the output file

def get_cpu_memory_usage():
    """Get CPU memory usage using adb shell command."""
    try:
        output = subprocess.check_output(["adb", "shell", "cat", "/proc/meminfo"])
        lines = output.decode("utf-8").splitlines()
        mem_total = None
        mem_free = None

        for line in lines:
            if "MemTotal" in line:
                mem_total = int(line.split()[1])
            elif "MemFree" in line:
                mem_free = int(line.split()[1])

        if mem_total is not None and mem_free is not None:
            used_ram = mem_total - mem_free
            free_memory_percentage = round((mem_free / mem_total) * 100, 2)
            return used_ram, free_memory_percentage, mem_total  # Return mem_total

    except Exception as e:
        print(f"Error getting CPU memory usage: {str(e)}")

    return 0, 0, 0


def memory_calc():
    print("Calculating CPU Memory Usage")
    cpu_memory_usages_no_app = []
    for i in range(10):
        used_ram, free_memory_percentage, mem_total = get_cpu_memory_usage()
        cpu_memory_usages_no_app.append((used_ram, free_memory_percentage, mem_total))
        time.sleep(1)

    # Calculate average values
    average_used_ram = sum(usage[0] for usage in cpu_memory_usages_no_app) / len(cpu_memory_usages_no_app)
    average_free_memory_percentage = sum(usage[1] for usage in cpu_memory_usages_no_app) / len(cpu_memory_usages_no_app)

    # Get total memory available (in bytes)
    _, _, total_memory_available_bytes = get_cpu_memory_usage()

    # Convert total memory from bytes to GB
    total_memory_available_gb = total_memory_available_bytes / (1024 * 1024 * 1024)

    # Calculate memory usage percentage
    memory_usage_perc = 100-(free_memory_percentage)

    # Print summary
    print("\nSummary:")
    print(f"Memory Used: {round(average_used_ram/1000, 1)} MB")
    print(f"Memory Used %: {memory_usage_perc}%")

    home_folder = os.path.expanduser("~")
    folder_path = os.path.join('/','Users',home_folder,'.pnp', 'traces')
    most_recent_folder = get_most_recent_folder(folder_path)
    if most_recent_folder is None:
        print("Error: Could not find most recent folder")
        return

    output_file = os.path.join('/','Users',home_folder,'.pnp', 'traces', most_recent_folder, 'artifacts', 'cpu_credit.csv')

    # Read the entire file into memory
    with open(output_file, 'r', newline='') as csvfile:
        reader = csv.reader(csvfile)
        rows = list(reader)

    # Modify the desired row
    rows[0].append(f"Memory Usage Percent")
    rows[0].append(f"{memory_usage_perc:.2f}")



    # Write the updated data back to the file
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(rows)



def extract_nugpu_avg_value():
    """
    Extracts the largest value of the key "nugpuAvg" from the gpu_stats_summary.json file.
    """
    print("Entering extract_nugpu_avg_value function")
    try:


        home_folder = os.path.expanduser("~")
        folder_path = os.path.join('/','Users',home_folder,'.pnp', 'traces')
        most_recent_folder = get_most_recent_folder(folder_path)
        if most_recent_folder is None:
            print("Error: Could not find most recent folder")
            return


        GPU_output_file = os.path.join('/','Users',home_folder,'.pnp', 'traces', most_recent_folder, 'artifacts', 'gpu_stats_summary.json')


        print("Trying to open gpu_stats_summary.json file")
        # Read the json file
        with open(GPU_output_file, 'r') as f:
            print("File opened successfully")
            data = json.load(f)
            print("JSON data loaded successfully")
            #print("JSON data:", data)


        # Recursive search for the key "nugpuAvg"
        def recursive_search(data):
            max_nugpu_avg = float('-inf')
            if isinstance(data, dict):
                for key, value in data.items():
                    if key == 'nugpuAvg':
                        max_nugpu_avg = max(max_nugpu_avg, value)
                    else:
                        result = recursive_search(value)
                        if result is not None:
                            max_nugpu_avg = max(max_nugpu_avg, result)
            elif isinstance(data, list):
                for item in data:
                    result = recursive_search(item)
                    if result is not None:
                        max_nugpu_avg = max(max_nugpu_avg, result)
            return max_nugpu_avg


        nugpu_avg_value = recursive_search(data)
        print("Recursive search complete")


        if nugpu_avg_value == float('-inf'):
            print("Key 'nugpuAvg' not found in the gpu_stats_summary.json file.")
        else:
            print("Key 'nugpuAvg' found with largest value:", nugpu_avg_value)

            output_file = os.path.join('/','Users',home_folder,'.pnp', 'traces', most_recent_folder, 'artifacts', 'cpu_credit.csv')

            # Read the entire file into memory
            with open(output_file, 'r', newline='') as csvfile:
                reader = csv.reader(csvfile)
                rows = list(reader)

            # Modify the desired row
            rows[1].append(f"NuGPUAverage")
            rows[1].append(f"{nugpu_avg_value:.2f}")

            # Write the updated data back to the file
            with open(output_file, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerows(rows)

    except FileNotFoundError:
        print("File 'gpu_stats_summary.json' not found.")
    except Exception as e:
        print("An error occurred:", str(e))


def get_most_recent_folder(folder_path):
    most_recent_folder = None
    most_recent_ctime = 0
    for dir in pathlib.Path(folder_path).iterdir():
        if dir.is_dir():
            ctime = dir.stat().st_ctime
            if ctime > most_recent_ctime:
                most_recent_ctime = ctime
                most_recent_folder = str(dir)
    return most_recent_folder



def os_system(cmd):
    if VERBOSE:
        print(f"Running command: {cmd}")
    subprocess.run(shlex.split(cmd))

def adb_cmd(cmd):
    if VERBOSE: # Hide the command unless verbose
        print(f"Running command: adb {cmd}")

    result = subprocess.run(["adb", cmd], capture_output=True, text=True)

    if VERBOSE and result.stdout != "": # print the output if verbose, hide it otherwise
        print(result.stdout.strip())

    # print any error messages
    if result.stderr != "":
        print(result.stderr)

    # return the resulting process if you need it.
    return result

def adb_shell(cmd):
    if VERBOSE: # Hide the command unless verbose
        print(f"Running command: adb shell {cmd}")

    result = subprocess.run(["adb", "shell", cmd], capture_output=True, text=True)

    if VERBOSE and result.stdout != "": # print the output if verbose, hide it otherwise
        print(result.stdout.strip())

    # print any error messages
    if result.stderr != "":
        print(result.stderr)

    # return the resulting process if you need it.
    return result

def find_pnp():
    # returns a path to pnp.par checks the known locations that it's likely to be installed on different computers (windows/mac feature install, and in the directory with this script)
    is_windows = sys.platform.startswith('win')

    if is_windows:
        #check if there is a chcolatey environment variable and use it if it exists.
        chocolatey = os.getenv('ChocolateyInstall')
        if chocolatey != None:
            pnp = Path(chocolatey).joinpath('lib','fb-pnp-windows','bin','pnp.par')
            if VERBOSE:
                print(f'Detected Windows system with chocolatey. Using pnp script: {pnp}')
        else:
            pnp = Path('C:\\','ProgramData','chocolatey','lib','fb-pnp-windows','bin','pnp.par')
            if VERBOSE:
                print(f'Detected Windows system. Using pnp script: {pnp}')
    else:
        #Mac
        pnp = Path('/','usr','local','bin','pnp')
        if VERBOSE:
            print(f'Detected MAC system: Using pnp script: {pnp}')

    # if pnp is not in one of the known install locations look for a local file, check script location first, then cwd then just fallback to raw 'pnp'
    if not pnp.exists():
        cwd_pnp = Path.cwd().joinpath("pnp.par")
        local_pnp = Path(sys.path[0]).joinpath("pnp.par")
        if VERBOSE:
                print('WARNING: pnp not found in default install locations.')
        if local_pnp.exists():
            pnp = local_pnp
            if VERBOSE:
                print('pnp.par found in script location')
        elif cwd_pnp.exists:
            pnp = cwd_pnp
            if VERBOSE:
                print('pnp.par found in current working directory')
        else:
            pnp = 'pnp'
            print('ERROR: Could not find pnp file, using fallback')
    if VERBOSE:
        print(f'Usiong PNP script: {pnp}')
    return pnp

def run_cpu_test(scenario, upload=False, local=False):


    is_windows = sys.platform.startswith('win')
    pnp = find_pnp()
    global PY_COMMAND

    #If override is not set, gets the path to the python installation that is running this script.
    if PY_COMMAND == None:
        PY_COMMAND = Path(sys.executable)
        if VERBOSE:
            print(f'PY_COMMAND automatically using value: {PY_COMMAND}')
    else:
        if VERBOSE:
            print(f'PY_COMMAND forced to use value: {PY_COMMAND}')

    if is_windows:
        current_dir = os.getcwd()
        filename = 'krishnav.txt'
        cmd = [pnp, '--note', 'perf_qa', '--report-cpu-credit', '--process-stats-poll-ms', '5000', '--duration', '60', '--scenario', str(scenario), "--enable-dsp-rpc-stats"]
    else:
        current_dir = os.getcwd()
        filename = 'krishnav.txt'
        if VERBOSE:
            print("Detected Mac environment, running pnp as an executable")
        cmd = [pnp, '--note', 'perf_qa', '--report-cpu-credit', '--process-stats-poll-ms', '5000', '--duration', '60', '--scenario', str(scenario), "--enable-dsp-rpc-stats"]


    if not upload and not local:
        user_confirmation = input("Is this test Local Only? (y,n):")
        if user_confirmation.strip().lower() in ["y","yep","yes","yeah"]:
            print("Will not upload this run.")
            cmd.append('--local-only')
    elif local:
        if VERBOSE:
            print("Running this test local-only")
        cmd.append('--local-only')
    else:
        print("Will upload this run.")

    if VERBOSE:
        print(f"Running command: {cmd}")

    if not TESTING:
        subprocess.run(cmd)

def set_oculuspreference(preferance, value):
    # takes two strings
    # Verifies that a given value is set for a given oculuspreference, sets it to the given value if it is not.
    # Returns True if it changes the value.

    result = adb_shell(f"oculuspreferences --get {preferance}")
    if f"[{preferance} : {value}] retrieved successfully" in result.stdout:
        if VERBOSE:
            print(f"{preferance} is already set to {value}")
        return False
    else:
        if VERBOSE:
            print(f"{preferance} is NOT set correctly! Setting to {value}...")
        result = adb_shell(f"oculuspreferences --set {preferance} {value}")
        if VERBOSE:
            print(result.stdout.strip())
        return True

def set_setprop(prop, value):
    # takes two strings
    # Verifies that a given value is set for a given oculuspreference, sets it to the given value if it is not.
    # Returns True if it changes the value.

    result = adb_shell(f"getprop {prop}")
    if value in result.stdout:
        if VERBOSE:
            print(f"{prop} is already set to {value}")
        return False
    else:
        if VERBOSE:
            print(f"{prop} is NOT set correctly! Setting to {value}...")
        result = adb_shell(f"setprop {prop} {value}")
        if VERBOSE:
            print(result.stdout.strip())
        return True

def set_oculussetting(setting, value):
    # takes two strings
    # Verifies that a given value is set for a given oculussetting, sets it to the given value if it is not.
    # Returns True if it changes the value.

    result = adb_shell(f"oculussetting --get {setting}")
    if f"{setting}={value}" in result.stdout:
        if VERBOSE:
            print(f"{setting} is already set to {value}")
        return False
    else:
        if VERBOSE:
            print(f"{setting} is NOT set correctly! Setting to {value}...")
        result = adb_shell(f"oculussetting --set {setting} {value}")
        if VERBOSE:
            print(result.stdout.strip())
        return True

def start_logcat(scenario):
    # Starts a logcat.
    # Returns the process running the logcat

    print("Starting Logcat")
    logcat_command = "adb logcat > BUC_" + str(scenario) + "_logcat.txt"
    logcat_process = subprocess.Popen(logcat_command, shell=True, start_new_session=True)
    return logcat_process

def stop_logcat(srm_process):
    # kill the logcat process
    adb_cmd("wait-for-device")
    srm_process.terminate()

def start_srm_logging():
    # Prepairs and starts SRM logging.
    # Requires adb root access and the SRM file in the same folder
    # Returns the process running the srm logging

    print("Pushing srm to /data/local/tmp")
    os_system("adb push srm /data/local/tmp")
    adb_shell("chmod +x data/local/tmp/srm")
    #adb_shell("/data/local/tmp/srm monitor")
    srm_process = subprocess.Popen("adb shell /data/local/tmp/srm monitor", shell=True, start_new_session=True)
    return srm_process

def get_srm_logs(srm_process, scenario):
    # zips and pulls the srm logs, then deletes the /tmp folder

    # kill the srm logging process
    adb_cmd("wait-for-device")
    srm_process.terminate()

    if VERBOSE:
        print("Compressing srm log files...")

    filename = "BUC_" + str(scenario) + "_srm_log.tar.gz"
    adb_shell(f"tar zcf data/local/tmp/{filename} data/local/tmp/srm_*.tsv")
    os_system(f'adb pull data/local/tmp/{filename}')
    adb_cmd("wait-for-device")

    # remove all the srm files and logs from the HMD
    adb_shell("rm data/local/tmp/*")
    print("\n================================================================================================\n")
    print(f"Pulled {filename}! Remember to rename or move the file or it will be overwritten next run!!")
    print("Don't forget to upload it to Drive!")
    print("\n================================================================================================\n")

def wait_for_background_tasks():
    # Tell the tester to wait before starting the trace in home scenarios

    print("After all the windows are open wait ~60 seconds for background CPU tasks to finish before starting the recording.")

def media_browser_overlay():
    # Setup for media browser overlay scenarios

    # set the preference and restart shell if needed
    #reboot = set_oculuspreference("debug_enable_auto_input_switching", "1")

    # if reboot:
    #    print("Restarting shell for settings to take effect...")
    #    adb_shell('am force-stop com.oculus.vrshell')
    #    adb_cmd('wait-for-device')
    #    time.sleep(10) # wait for shell to restart
    print("Turn on Seemless multitasking under Settings > Experimantal Settings")


def first_encounters():
    # Setup for first encounters scenarios


    # adb shell oculussetting --set hand_tracking_opt_in 1
    # adb shell setprop debug.oculus.bodyApiForceFullBody 1
    # adb shell setprop debug.oculus.bodyApiCmd keepalive
    # adb shell setprop debug.oculus.forceEnablePerformanceFeatures "BODY_TRACKING"
    # adb shell setprop debug.oculus.forceHighFidelityPerformanceFeatures "BODY_TRACKING"
    # adb shell setprop debug.oculus.forceMediumFidelityPerformanceFeatures null
    # adb shell am start-foreground-service com.oculus.bodyapiservice/.BodyApiService

    set_oculussetting("hand_tracking_opt_in", "1")
    set_setprop("debug.oculus.bodyApiForceFullBody", "1")
    set_setprop("debug.oculus.bodyApiCmd", "keepalive")
    set_setprop("debug.oculus.forceEnablePerformanceFeatures", "BODY_TRACKING")
    set_setprop("debug.oculus.forceHighFidelityPerformanceFeatures", "BODY_TRACKING")
    set_setprop("debug.oculus.forceMediumFidelityPerformanceFeatures", "null")
    adb_shell("am start-foreground-service com.oculus.bodyapiservice/.BodyApiService")

def enable_body_tracking():
    set_oculussetting("hand_tracking_opt_in", "1")
    set_setprop("debug.oculus.bodyApiForceFullBody", "1")
    set_setprop("debug.oculus.bodyApiCmd", "keepalive")
    set_setprop("debug.oculus.forceEnablePerformanceFeatures", "BODY_TRACKING")
    set_setprop("debug.oculus.forceHighFidelityPerformanceFeatures", "BODY_TRACKING")
    set_setprop("debug.oculus.forceMediumFidelityPerformanceFeatures", "null")
    adb_shell("am start-foreground-service com.oculus.bodyapiservice/.BodyApiService")
    print("Body Tracking Enabled")


def enable_hand_tracking():
    set_oculussetting("hand_tracking_opt_in", "1")
    print("Hand Tracking Enabled")

def disable_hand_tracking():
    # disable handtracking in scenarios that don't use it
    if VERBOSE:
        print("Disabling hand tracking.")

    set_oculussetting("hand_tracking_opt_in", "0")

def open_youtube():
    cmd1 = "adb shell am force-stop com.oculus.browser"
    cmd2 = "adb shell pm clear com.oculus.browser"
    cmd3 = "adb shell am broadcast -a com.oculus.vrshell.intent.action.LAUNCH -n com.oculus.vrshell/.ShellControlBroadcastReceiver -d apk://com.oculus.browser -e uri https://www.youtube.com/watch?v=RzVvThhjAKw&quality=hd720"
    subprocess.run(cmd1, shell=True)
    subprocess.run(cmd2, shell=True)
    subprocess.run(cmd3, shell=True)

    print("Opened Youtube (Browser)")
    #https://www.youtube.com/watch?v=RzVvThhjAKw&vq=hd720
    time.sleep(1)


def close_apps():
    cmd1 = "adb shell am force-stop com.oculus.socialplatform"
    cmd2 = "adb shell am force-stop com.oculus.explore"

    subprocess.run(cmd1, shell=True)
    subprocess.run(cmd2, shell=True)

    print("Closed Startup apps")

def restart_shell():
    subprocess.run(['adb', 'shell', 'am', 'force-stop', 'com.oculus.vrshell'])

def open_startup_apps():
    os.system("adb shell am start -n com.oculus.vrshell/.MainActivity -d com.oculus.socialplatform")
    time.sleep(1)
    os.system("adb shell am start -n com.oculus.vrshell/.MainActivity -d com.oculus.explore")
    time.sleep(1)




def open_heavy_tasks():
    os.system("adb shell am start com.oculus.igvr")
    time.sleep(1)
    os.system("adb shell am start -n com.whatsapp/com.whatsapp.Main")
    time.sleep(1)
    os.system("adb shell am start com.oculus.remotedesktop")
    time.sleep(1)



def spotify():
    print("Run Spotify PWA in the background")

def horizon_worlds():
    print('Run "Horizon Worlds", play "Super Rumble World"')

def casting():
    print('Start casting to your mobile device')

def quest_call():
    print("-----------------------------------------------------------------------------------")
    print("Start a call with seven bots:")
    print("     Go to: https://www.internalfb.com/intern/oculus/parties/echo_tester")
    print('     Set Number of test users to "7", then press "invite to play"')
    print('     On your HMD join the incoming call')
    print("-----------------------------------------------------------------------------------")

def remote_display():
    print('Launch: "Remote Display (Beta)"')
    print('Connect to your PC')

def six_panel():
    print("-----------------------------------------------------------------------------------")
    print('Enable 6 spacial panels:')
    #print('    Launch "Shell Debug Panel" from Internal Settings -> Launch "Shell Debug Settings"')
    #print('    In the  “Spatial Panels Experience” tab:')
    #print('    Set “Spacial Panels Experience” to "2" and “Panels in Primary Layout” to "6"')
    #print('    Restart Shell by "Render Preferences" > "Restart Shell"')
    print('    NOTE(4/25/24): To get more than three panels you must drag panels to be free floating')
    print('    NOTE(8/20/24): 6 spacial panels should be enabled by default on all builds')
    print("-----------------------------------------------------------------------------------")

def enableVR():

    cmd1 = "adb root"
    cmd2 = "adb shell oculuspreferences --setc reality_tuner_value 0 && adb shell am broadcast -a com.oculus.vrguardianservice.JsonCmdUserBroadcast --es cmd '{\"requestPassthrough\":{\"status\":1,\"method\":3}}'"
    subprocess.run(cmd1, shell=True)
    subprocess.run(cmd2, shell=True)

    print("VR Enabled")


def enableMR():
    cmd1 = "adb root"
    cmd2 = "adb shell oculuspreferences --setc reality_tuner_value 100 && adb shell am broadcast -a com.oculus.vrguardianservice.JsonCmdUserBroadcast --es cmd '{\"requestPassthrough\":{\"status\":2,\"method\":3}}'"

    subprocess.run(cmd1, shell=True)

    #os.system("adb shell am force-stop com.oculus.vrshell")

    subprocess.run(cmd2, shell=True)

    print("MR Enabled")




def stayAwake():
    print("Staying Awake")
    os.system("adb shell svc power stayon true")
    os.system("adb shell setprop persist.oculus.travel_mode 1")
    os.system("adb shell am broadcast -a com.oculus.vrpowermanager.prox_close")
    os.system("adb shell setprop persist.oculus.guardian_disable 1")




def simulate_enter_key_press():
    # Wait for 5 seconds
    time.sleep(30)

    # Simulate Enter key press
    print("Pressing Enter")
    pyautogui.press('enter')


def default_scenario():
    print("Default Scenario")

def all_scenarios():
    if VERBOSE:
        print("python location: " + os.path.dirname(sys.executable))

    print("Waiting for device...")
    adb_cmd('wait-for-device')
    adb_cmd('root')
    adb_cmd('wait-for-device')
    # Stop telemetry
    #adb_shell('stop telemetry')
    # disable power hints
    if not TESTING:
        adb_shell('lshal debug android.hardware.power@1.3::IPower/default disable-hint')

def post_test_cleanup():
    # Cleanup commands to reset the hmd
    print("Cleaning up after test run.")
    # Reset hand tracking override
    #subprocess.check_output("adb shell oculuspreferences --set hand_tracking_override_frequency 0",stderr=subprocess.STDOUT,shell=True)
    time.sleep(2)

    extract_nugpu_avg_value()


    reboot = input("Press any button to reboot or 'n' to skip...")

    if reboot == 'n':
        return

    # Reboot the device
    if not TESTING:
        adb_cmd('reboot')

def scenario_4():
    print("Scenario 4: Beatsaber")
    print("================================================================================================")
    disable_hand_tracking()
    print("Solo -> OST Vol. 1 -> Breezer -> Expert Plus")
    print("Standard Controls, No Fail Mode")
    print("Start collecting data when the level starts")

def scenario_7():
    print("Scenario 7: Population One + Casting")
    print("================================================================================================")
    disable_hand_tracking()
    casting()
    print("Training -> Single Player Battle Royal -> Play")
    print("Start capturing data when the level starts")
    print('When the level starts: Run to the launcher, wait for the fuel level to reach "30" then pull the handle')
    print('When you land, pull out the knife, then press B and paint the glasses nearby')

def scenario_8():
    print("Scenario 8: Workrooms Solo Meeting")
    print("================================================================================================")
    #Enable Lip Sync
    flicker = set_setprop("debug.oculus.ft.forceEnable", "1")
    if flicker:
        subprocess.check_output("adb shell input keyevent 26",stderr=subprocess.STDOUT,shell=True)
        time.sleep(2)
        subprocess.check_output("adb shell input keyevent 26",stderr=subprocess.STDOUT,shell=True)

    #Enable hands
    set_oculussetting("hand_tracking_opt_in", "1")

    print("Launch Meta Horizon Workrooms (Beta)")
    print("Stay at the solo desk and enable the passthrough environment.")
    print("Connect to remote desktop, on your PC start a Youtube video and have a document open.")
    print("Set down the controllers before starting the recording. During the recording keep your hands in view and moving")

def scenario_12():
    print("Scenario 12: Demeo (MR gaming)")
    print("================================================================================================")
    disable_hand_tracking()
    print("Turn on MR mode. Settings -> AR")
    print("Skirmish -> New Game -> Black Sarcophagus -> Sorcerer")

def scenario_101():
    print("Scenario 101: VR Content + Spotify")
    print("================================================================================================")
    disable_hand_tracking()
    spotify()
    horizon_worlds()


def scenario_102():
    print("Scenario 102: VR Content + Casting")
    print("================================================================================================")
    disable_hand_tracking()
    casting()
    horizon_worlds()

def scenario_103():
    print("Scenario 103: VR Content + Twitch in Browser in Overlay")
    print("================================================================================================")
    disable_hand_tracking()
    media_browser_overlay()
    print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    print("Make sure the browser window is visible over Horizon Worlds.")
    horizon_worlds()

def scenario_104():
    print("Scenario 104: VR Content + Heavy Multitasking (Spotify, Casting, Twitch in Browser)")
    print("================================================================================================")
    disable_hand_tracking()
    spotify()
    casting()
    media_browser_overlay()
    print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    print("Make sure the browser window is visible over Horizon Worlds.")
    horizon_worlds()

def scenario_105():
    print("Scenario 105:  MR Content + Spotify")
    print("================================================================================================")
    spotify()
    first_encounters()
    print('Launch "First Encounters"')
    print("Start recording after the aliens enter your room.")

def scenario_106():
    print("Scenario 106:  MR Content + Casting")
    print("================================================================================================")
    casting()
    first_encounters()
    print('Launch "First Encounters"')
    print("Start recording after the aliens enter your room.")

def scenario_107():
    print("Scenario 107:  MR Content + Pinned Overlay")
    print("================================================================================================")
    media_browser_overlay()
    first_encounters()
    print('Launch "First Encounters"')
    print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    print("Make sure the browser window is visible over First Encounters.")
    print("Start recording after the aliens enter your room.")

def scenario_108():
    print("Scenario 108:  MR Content + Heavy Multitasking")
    print("================================================================================================")
    media_browser_overlay()
    first_encounters()
    spotify()
    print('Launch "First Encounters"')
    casting()
    print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    print("Make sure the browser window is visible over First Encounters.")
    print("Start recording after the aliens enter your room.")

def scenario_110():
    print("Scenario 110: VR Home - Sports Fan")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in VR home "Meta Horizon Terrace" environment')
    print('Use nearfield “tablet” mode vr shell')
    quest_call()
    #print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    #print("Open a second browser window, drag it next to the exiting one, go to youtube.com and start streaming a video. Make sure the video is not muted.")
    #print("Open a third browser window, drag it next to the exiting one, go to youtube.com and start streaming a video. Make sure the video is not muted.")
    print("Open THREE browser windows: ")
    print("    1. Twitch.com")
    print("    2. Youtube.com")
    print("    3. Youtube.com")
    print("Make sure all streams and videos are playing and not muted.")
    wait_for_background_tasks()

def scenario_111():
    print("Scenario 111: VR Home - Solo Productivity")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in VR home "Meta Horizon Terrace" environment')
    print('Use nearfield “tablet” mode vr shell')
    spotify()
    #print("Open one browser window, go to docs.google.com.")
    #print("Open a second browser window, drag it next to the exiting one, go to sheets.google.com")
    print("Open TWO browser windows: ")
    print("    1. docs.google.com")
    print("    2. sheets.google.com")
    remote_display()
    wait_for_background_tasks()
    print('While recording, make sure to continuously interact by scrolling in the Google Doc window.')

def scenario_112():
    print("Scenario 112: VR Home - 6-Screen Torture Test")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in VR home "Meta Horizon Terrace" environment')
    print('Use nearfield “tablet” mode vr shell')
    six_panel()
    casting()
    remote_display()
    #print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    #print("Open a second browser window, drag it next to the exiting one, go to twitch.com and start streaming a popular stream. Make sure the video is not muted.")
    #print("Open a third browser window, drag it next to the exiting one, go to twitch.com and start streaming a popular stream. Make sure the video is not muted.")
    print("Open THREE browser windows: ")
    print("    1. Twitch.com")
    print("    2. Twitch.com")
    print("    3. Twitch.com")
    print("Make sure all streams are playing and not muted.")
    print("Launch Store Panel App")
    print("Launch Horizon Feed panel App")
    wait_for_background_tasks()

def scenario_114():
    print("Scenario 114: MR Home - Sports Fan")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in MR passthrough home')
    print('Use nearfield “tablet” mode vr shell')
    #print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    #print("Open a second browser window, drag it next to the exiting one, go to youtube.com and start streaming a video. Make sure the video is not muted.")
    #print("Open a third browser window, drag it next to the exiting one, go to youtube.com and start streaming a video. Make sure the video is not muted.")
    print("Open THREE browser windows: ")
    print("    1. Twitch.com")
    print("    2. Youtube.com")
    print("    3. Youtube.com")
    print("Make sure all streams and videos are playing and not muted.")
    quest_call()
    wait_for_background_tasks()

def scenario_115():
    print("Scenario 115: MR Home - Solo Productivity")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in MR passthrough home')
    print('Use nearfield “tablet” mode vr shell')
    spotify()
    #print("Open one browser window, go to docs.google.com.")
    #print("Open a second browser window, drag it next to the exiting one, go to sheets.google.com")
    print("Open TWO browser windows: ")
    print("    1. docs.google.com")
    print("    2. sheets.google.com")
    remote_display()
    wait_for_background_tasks()
    print('While recording, make sure to continuously interact by scrolling in the Google Doc window.')

def scenario_116():
    print("Scenario 116: MR Home - 6-Screen Heavy Multitasking")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in MR passthrough home')
    print('Use nearfield “tablet” mode vr shell')
    six_panel()
    casting()
    remote_display()
    #print("Open one browser window, go to twitch.com and start streaming a popular stream. Make sure the stream is not muted.")
    #print("Open a second browser window, drag it next to the exiting one, go to twitch.com and start streaming a popular stream. Make sure the video is not muted.")
    #print("Open a third browser window, drag it next to the exiting one, go to twitch.com and start streaming a popular stream. Make sure the video is not muted.")
    print("Open THREE browser windows: ")
    print("    1. Twitch.com")
    print("    2. Twitch.com")
    print("    3. Twitch.com")
    print("Make sure all streams are playing and not muted.")
    print("Launch Store Panel App")
    print("Launch Horizon Feed panel App")
    wait_for_background_tasks()

def scenario_117():
    print("Scenario 117: VR Home - 6-Screen Light Multitasking")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in VR home "Meta Horizon Terrace" environment')
    print('Use nearfield “tablet” mode vr shell')
    six_panel()
    #print("Open one browser window, go to google.com.")
    #print("Open a second browser window, drag it next to the exiting one, go to google.com.")
    #print("Open a third browser window, drag it next to the exiting one, go to go to google.com.")
    print("Open THREE browser windows: ")
    print("    1. google.com")
    print("    2. google.com")
    print("    3. google.com")
    print("Launch People Panel App")
    print("Launch Store panel App")
    print("Launch Notifications Panel App")
    wait_for_background_tasks()

def scenario_118():
    print("Scenario 118: MR Home - 6-Screen Light Multitasking")
    print("================================================================================================")
    disable_hand_tracking()
    print('While in MR passthrough home')
    print('Use nearfield “tablet” mode vr shell')
    six_panel()
    #print("Open one browser window, go to google.com.")
    #print("Open a second browser window, drag it next to the exiting one, go to google.com.")
    #print("Open a third browser window, drag it next to the exiting one, go to go to google.com.")
    print("Open THREE browser windows: ")
    print("    1. google.com")
    print("    2. google.com")
    print("    3. google.com")
    print("Launch People Panel App")
    print("Launch Store panel App")
    print("Launch Notifications Panel App")
    wait_for_background_tasks()

def scenario_301():
    print("Scenario 301: VR Home + AUI Only")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()



def scenario_302():
    print("Scenario 302: VR 2 Person Co-Presence + Youtube (Browser)")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()

def scenario_303():
    print("Scenario 303: VR 2 Person Co-Presence Heavy")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_304():
    print("Scenario 304: VR 2 Person Co-Presence + Heavy + Nav")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_305():
    print("Scenario 305: VR 4 Person Co-Presence + Youtube (Browser)")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()


def scenario_306():
    print("Scenario 306: VR 4 Person Co-Presence + Heavy")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()



def scenario_307():
    print("Scenario 307: VR 4 Person Co-Presence + Heavy + Nav")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_308():
    print("Scenario 308: VR 8 Person Co-Presence + Youtube (Browser)")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()


def scenario_309():
    print("Scenario 309: VR 8 Person Co-Presence + Heavy")
    print("================================================================================================")
    enableVR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_310():
    print("Scenario 310: MR Home + AUI Only")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()

def scenario_311():
    print("Scenario 311: MR 2 Person Co-Presence + Youtube (Browser)")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()

def scenario_312():
    print("Scenario 312: MR 2 Person Co-Presence + Heavy")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_313():
    print("Scenario 313: MR 2 Person Co-Presence + Heavy + Nav")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_314():
    print("Scenario 314: MR 4 Person Co-Presence + Youtube (Browser)")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()

def scenario_315():
    print("Scenario 112: MR 4 Person Co-Presence + Heavy")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_316():
    print("Scenario 316: MR 4 Person Co-Presence + Heavy + Nav")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_317():
    print("Scenario 317: MR 8 Person Co-Presence + Youtube (Browser)")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()

def scenario_318():
    print("Scenario 318: MR 8 Person Co-Presence + Heavy")
    print("================================================================================================")
    enableMR()
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_401():
    print("Scenario 401: VR Immersive App Baseline")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()

def scenario_402():
    print("Scenario 402: VR Immersive App + Youtube")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()

def scenario_403():
    print("Scenario 403: VR Immersive App + Youtube + People + Horizon Feed")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()

def scenario_404():
    print("Scenario 404: VR Immersive App + Heavy")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()

def scenario_405():
    print("Scenario 405: MR Immersive App Baseline")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()

def scenario_406():
    print("Scenario 406: MR Immersive App + Youtube")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_youtube()

def scenario_407():
    print("Scenario 407: MR Immersive App + Youtube + People + Horizon Feed")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()

def scenario_408():
    print("Scenario 408: MR Immersive App + Heavy")
    print("================================================================================================")
    enable_hand_tracking()
    enable_body_tracking()
    close_apps()
    open_startup_apps()
    open_youtube()
    print("================================================================================================")
    print("Move these apps up before continuing")
    input("Press any button to continue...")
    open_heavy_tasks()




def main():
    # gets arguments from the command line
    parser = argparse.ArgumentParser(description="Sets up PnP test scenarios and launches the collection script.")
    # Required argument of a filepath to work on
    parser.add_argument("-s", "--scenario", dest='scenario', help="The Scenario to be run.")
    parser.add_argument("-u", "--upload", dest="upload", action='store_true', default=False, help="Upload the test run. Takes precidence over -l")
    parser.add_argument("-l", "--local_only", dest="local", action='store_false', default=True, help="Run the test Local-Only and don't ask to upload.")
    parser.add_argument("-d", "--default", dest="default_test", action='store_true', default=False, help="Run as a default test. Will only run common pre-test commands. Will not upload.")
    parser.add_argument("-v", "--verbose", dest="verbose", action='store_true', default=False, help="Prints more command line feedback.")
    parser.add_argument("--srm", dest="srm", action='store_true', default=False, help="Enables additional SRM logging.")
    parser.add_argument("--logcat", dest="logcat", action='store_true', default=False, help="Enables additional SRM logging.")
    args = parser.parse_args()


    # set global VERBOSE flag
    global VERBOSE
    VERBOSE = args.verbose



    # Get the path of the current file
    current_file_path = __file__
    #Get the directory of the current file
    current_dir = os.path.dirname(current_file_path)
    # Specify the name of the file you're looking for
    target_file_name = 'BUC_Config.txt'
    # Construct the full path of the target file
    target_file_path = os.path.join(current_dir, target_file_name)
    # Check if the target file exists
    if os.path.exists(target_file_path):
        print(f"Found {target_file_name} in {current_dir}")
        config=target_file_path
    else:
        print(f"{target_file_name} not found in {current_dir}")







    # if this test if for HCL, or a default test, force it to be local
    if args.default_test == True:
        local = True
        upload = False
        if VERBOSE:
            print(f"Forcing test to be local_only because default_test: {args.default_test}")
    else:
        # If
        if args.upload == True:
            local = False
            upload = True
        else:
            local = args.local
            upload = args.upload


    if args.scenario == None:
        user_input = input("Enter Scenario #:")
        if user_input == "":
            scenario = 0
        else:
            scenario = int(user_input)
    else:
        scenario = int(args.scenario)


    # run prep commands for all scenarios
    all_scenarios()


    # start SRM logging if requested
    if args.srm == True:
        srm_process = start_srm_logging()


    # start a logcat if requested
    if args.logcat == True:
        logcat_process = start_logcat(scenario)


    if args.default_test == True:
        default_scenario()


    elif scenario == 4:
        scenario_4()


    elif scenario == 7:
        scenario_7()


    elif scenario == 8:
        scenario_8()


    elif scenario == 12:
        scenario_12()


    elif scenario == 101:
        scenario_101()


    elif scenario == 102:
        scenario_102()


    elif scenario == 103:
        scenario_103()


    elif scenario == 104:
        scenario_104()


    elif scenario == 105:
        scenario_105()


    elif scenario == 106:
        scenario_106()


    elif scenario == 107:
        scenario_107()


    elif scenario == 108:
        scenario_108()


    elif scenario == 110:
        scenario_110()


    elif scenario == 111:
        scenario_111()


    elif scenario == 112:
        scenario_112()


    elif scenario == 114:
        scenario_114()


    elif scenario == 115:
        scenario_115()


    elif scenario == 116:
        scenario_116()


    elif scenario == 117:
        scenario_117()


    elif scenario == 118:
        scenario_118()


    elif scenario == 301:
        scenario_301()


    elif scenario == 302:
        scenario_302()


    elif scenario == 303:
        scenario_303()


    elif scenario == 304:
        scenario_304()


    elif scenario == 305:
        scenario_305()

    elif scenario == 306:
        scenario_306()

    elif scenario == 307:
        scenario_307()

    elif scenario == 308:
        scenario_308()

    elif scenario == 309:
        scenario_309()

    elif scenario == 310:
        scenario_310()

    elif scenario == 311:
        scenario_311()

    elif scenario == 312:
        scenario_312()

    elif scenario == 313:
        scenario_313()

    elif scenario == 314:
        scenario_314()

    elif scenario == 315:
        scenario_315()

    elif scenario == 316:
        scenario_316()

    elif scenario == 317:
        scenario_317()

    elif scenario == 318:
        scenario_318()

    elif scenario == 401:
        scenario_401()

    elif scenario == 402:
        scenario_402()

    elif scenario == 403:
        scenario_403()

    elif scenario == 404:
        scenario_404()

    elif scenario == 405:
        scenario_405()

    elif scenario == 406:
        scenario_406()

    elif scenario == 407:
        scenario_407()

    elif scenario == 408:
        scenario_408()

    elif scenario == 1301:
        for i in range(3):
            print("Scenario 1301: VR Home + AUI Only AUTOMATED")
            print("================================================================================================")
            os.system("adb shell input keyevent KEYCODE_WAKEUP")
            enableVR()
            stayAwake()
            enable_hand_tracking()
            enable_body_tracking()
            close_apps()
            print("starting threading")
            timer_thread = threading.Thread(target=simulate_enter_key_press)
            timer_thread.start()
            print("starting pnp")
            run_cpu_test(scenario,upload, local)
            time.sleep(20)
            print(f"Test {i+1} Complete, Rebooting")
            time.sleep(2)
            if not TESTING:
                adb_cmd('reboot')
            time.sleep(30)
            stayAwake()

    elif scenario == 1302:
        for i in range(3):
            print("Scenario 1302: VR Home + youtube AUTOMATED")
            print("================================================================================================")
            os.system("adb shell input keyevent KEYCODE_WAKEUP")
            enableVR()
            stayAwake()
            enable_hand_tracking()
            enable_body_tracking()
            close_apps()
            open_youtube()
            time.sleep(5)
            os.system("adb shell input text K")
            print("starting threading")
            timer_thread = threading.Thread(target=simulate_enter_key_press)
            timer_thread.start()
            print("starting pnp")
            run_cpu_test(scenario, upload, local)
            time.sleep(20)
            print(f"Test {i+1} Complete, Rebooting")
            time.sleep(2)
            if not TESTING:
                adb_cmd('reboot')
            time.sleep(30)
            stayAwake()


    elif scenario == 1310:
        for i in range(3):
            print("Scenario 1310: MR Home + AUI Only AUTOMATED")
            print("================================================================================================")
            os.system("adb shell input keyevent KEYCODE_WAKEUP")
            #os.system("adb reboot")
            #time.sleep(60)
            os.system("adb shell setprop persist.oculus.travel_mode 0")
            os.system("adb shell setprop persist.oculus.guardian_disable 0")
            enableMR()
            stayAwake()

            enable_hand_tracking()
            enable_body_tracking()
            close_apps()
            print("starting threading")
            timer_thread = threading.Thread(target=simulate_enter_key_press)
            timer_thread.start()
            print("starting pnp")
            run_cpu_test(scenario ,upload, local)
            time.sleep(20)
            print(f"Test {i+1} Complete, Rebooting")
            time.sleep(2)
            if not TESTING:
                adb_cmd('reboot')
            time.sleep(30)
            stayAwake()

    elif scenario == 1311:
        for i in range(3):
            print("Scenario 1311: MR Home + youtube AUTOMATED")
            print("================================================================================================")
            os.system("adb shell input keyevent KEYCODE_WAKEUP")
            enableMR()
            os.system("adb shell setprop persist.oculus.travel_mode 0")
            os.system("adb shell setprop persist.oculus.guardian_disable 0")
            os.system("adb reboot")
            time.sleep(30)
            stayAwake()
            enable_hand_tracking()
            enable_body_tracking()
            close_apps()
            open_youtube()
            time.sleep(5)
            os.system("adb shell input text K")
            print("starting threading")
            timer_thread = threading.Thread(target=simulate_enter_key_press)
            timer_thread.start()
            print("starting pnp")
            run_cpu_test(scenario ,upload, local)
            time.sleep(20)
            print(f"Test {i+1} Complete, Rebooting")
            time.sleep(2)
            if not TESTING:
                adb_cmd('reboot')
            time.sleep(30)
            stayAwake()


    else:
        assert True, "Unkown scenario! Please specify a known scenario!!"

    # Run the PnP test script
    print("================================================================================================")

    if scenario not in (1301, 1302, 13010):

        print("starting pnp")
        run_cpu_test(scenario ,upload, local)


    memory_calc()


    # Get SRM logs if they were collected
    if args.srm == True:
        get_srm_logs(srm_process, scenario)

    # Stop logcat if they were running
    if args.srm == True:
        stop_logcat(logcat_process)


    # Run post test cleanup commands
    if scenario != 1301:
        post_test_cleanup()



if __name__ == "__main__":
    main()
