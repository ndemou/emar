emar, Easy MAnagment of Remote tasks
------------------------------------

emar helps you run powershell function(s) on many windows PCs/Servers 
in parallel and get back the results. It relies on PowerShell remoting to run the functions. 
You can think of emar as a glorified wrapper around Invoke-Command that provides these extra features:

 - Detect online clients (responding to ping) and only attempt tasks on them 
    
 - Periodicly retry failed clients.
    
 - Collect errors & results.
    
 - Keep detailed logs and status reports.
    
 - Easily run more than one tasks.
    
Known limitations

 - You can not return huge amounts of data because all the data from the parallel runs of tasks are cached in the servers memory.
 
 - The output of the stream must end with the string `"<SUCCESS>"` which is anoying if you mainly want to receive PS objects from the task (but is essential if you want to know if a task succeeded or not).

Getting started
---------------

 0) On all the clients you should have enabled PowerShell remoting. Running
    `EnterPSSession "Client-Computer-Name"`
    from the server should print the clients computer name
    
 All other steps are on the server.
    
 1) Create a directory for emar to work in (`$base_dir`).

        $base_dir="c:\it\emar" 
        mkdir -force $base_dir\tasks

 2) Select an id for your first task (`$task_id`) and create a folder for it 
    `mkdir $base_dir\$task_id`
    I use this style `'202010_Inst_Chrome'` (for 2020-Octomber, install chrome)
    The id can be anything that's a valid variable name but don't start it with _
    
        mkdir $base_dir\tasks\202011_Inst_Chrome

 3) Write a function and put it in `$base_dir\$task_id\task.ps1`
    The last (and maybe only) thing your function should return is the text 
    `"<SUCCESS>"` if it's job was done succesfully - anything else if not.
    You can write code to collect data from the clients or to perform 
    jobs like installing software.
    
    It's probably good to abort on any error.
    
    It's also a good idea to return clixml or json.
    
    It's a bad idea to return huge amounts of data (they are collected 
    in memory from all clients before getting written to disk)

        notepad $base_dir\tasks\202011.testemar\task.ps1
        #------------------------------------------------------------
        # How much time your function needs to complete 
        # (worst-case scenario )
        $script:TIMEOUT = 300

        function ClientTask() {
            # stop on any error
            $ErrorActionPreference = "Stop" 
            # Invoke-webrequest and others will not display progress
            $ProgressPreference = 'SilentlyContinue'    
            ...YOUR CODE HERE...
            if ($check_if_all_good) {
                echo "<SUCCESS>"
            } else {
                echo "HUSTON WE'VE HAD A PROBLEM"
            }    
        }
        #------------------------------------------------------------

 4) create a text file `$base_dir\clients.txt` with a list of computer names 
    (one per line) were you want to run your task on
    
        echo 'test-pc' > $base_dir\tasks\202011.testemar\clients.txt

 5) Execute emar

        emar.ps1 -command start -base_dir $base_dir

 6) As tasks run on clients:
 
    Output of sucessful tasks is saved in:
    
        $base_dir\tasks\$task_id\results\$computer_name.txt
        
    Outpute of failed tasks in:
    
        $base_dir\tasks\$task_id\bad.results.$computer_name.txt
        
    A nice summary of the current status is in:
    
        $base_dir\tasks\$task_id\status.txt
        
    Detailed logs are written in:
    
        $base_dir\tasks\$task_id\log.txt
        
    Status messages are printed on screen

An ugly outlook of emar
-----------------------

![Almost everything that matters in one image](emar-ugly-outlook.png)
