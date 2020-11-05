<#
EMAR, Easy MAnagment of Remote tasks
------------------------------------

emar helps you run a powershell function on many client-PCs and get back
results (if any). Your function may do anything powershell can do except
return back huge amounts of data. All in all emar is a glorified wrapper 
around Invoke-Command with extra logic to:
    - Detect and only attempt clients that respond to ping 
    - Periodic retry of failed clients
    - Collection of the results of successful tasks in files
    - Nice logs and reports (more lines of code than I thought)
    - Easily run more than one tasks

Getting started
---------------

 1) Create a directory for emar to work in ($base_dir).

        $base_dir="c:\it\emar" 
        mkdir -force $base_dir

 2) Select an id for your first task ($task_id) and mkdir $base_dir\$task_id
    I use this style '202010_Inst_Chrome' (for 2020-Octomber, install chrome)
    The id can be anything you like but don't start it with _
    (and no, don't use spaces -- use something like a variable name).
    
        mkdir $base_dir\202011.testemar

 3) Write a function and put it in $base_dir\$task_id\task.ps1
    The last (and maybe only) thing your function should return is the text 
    <SUCCESS> if it's job was done succesfully - anything else if not.
    You can write code to collect data from the clients or to perform 
    jobs like installing software.
    It's probably good to abort on any error.
    It's also a good idea to return clixml or json.
    It's a bad idea to return huge amounts of data (they are collected
    in memory from all clients before getting written to disk)

        notepad $base_dir\202011.testemar\task.ps1
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

 4) create a text file $base_dir\clients.txt with a list of computer names 
    (one per line) were you want to run your task on
    
        echo 'test-pc' > $base_dir\202011.testemar 

 5) Execute emar

        emar.ps1 -command start -base_dir $base_dir

 6) As tasks run on clients:
    Output of sucessful tasks is saved in:
        $base_dir\$task_id\results\<computer_name>.txt
    Outpute of unsuccesful tasks in:
        $base_dir\$task_id\bad.results.<computer_name>.txt
    A nice summary of the current status is in:
        $base_dir\$task_id\status.txt
    Detailed logs are written in:
        $base_dir\$task_id\log.txt
    Status messages with colors that should please the LGBTQ+ community
    are printed on screen

TODO
----
1. A variable in task.ps1 that will instruct emar to strip the last line
   of the output if it's <SUCCESS> so that I can write clean clixml or json
   e.g.:
     $script:STRIP_SUCCESS_FROM_LAST_LINE = $true

2. Write a completely different tool which will run on clients to periodicaly 
   poll a central server for tasks to execute and send back the results.
    - Everything must be signed (and maybe encrypted) to avoid security disaster
    - Show Extra care on how the code will auto-update itself without failing 
      and leaving the client without the ability to receive new jobs.
#>
param (
    [Parameter(Mandatory=$true)] [String]$command,
    [string]$Task_Id="",
    [string]$base_dir = "C:\it\emar"
)

$MAX_TASK_RETRIES=50
function log($msg, $color, $dont_print) {
    $ts = "{0:MM/dd} {0:HH:mm:ss}" -f (Get-Date)
    if (!($color)) {$color="Gray"}
    if ($dont_print) {} else {
    write-host "$ts $task_id $msg" -ForegroundColor $color}
    $prefix = ("$ts {0,11}" -f $color )
    "$prefix $msg" >> "$task_dir\log.txt"
}
function list_clients_in_one_line($clients_list) {
    # returns a nice short and sorted string with all clients in one line
    $temp = $clients_list | Sort-Object | ForEach-Object { $_ -replace 'RPS','' -replace '-PC',''}
    "$temp"
}
function date_time_to_str($dt) {
    "{0:yyyy-MM-dd} {0:HH:mm:ss}" -f ($dt)
}
function load_state($file) {
    # returns the $state deserialized from $file

    # $temp will be a psobject with properties (NOT a dictionary)
    $temp = (Get-Content $file | ConvertFrom-Json)

    # let's convert it to a dictionary
    $s=@{}
    $temp.PSObject.Properties | ForEach-Object { $s[$_.Name]=$_.Value }

    # fix the times (convert from string to [datetime])
    if ($s.deployment_start) {$s.deployment_start = [datetime]::parse($s.deployment_start)}
    if ($s.last_success) {$s.last_success = [datetime]::parse($s.last_success)}
    return $s
}
function save_state($state, $file) {
    # serializes and saves $state to $file
    $s = $state.clone()
    # convert times from [datetime] to string
    $s['deployment_start'] = (date_time_to_str $s['deployment_start'])
    $s['last_success'] = (date_time_to_str $s['last_success'])
    # write
    $s | ConvertTo-Json > $file
}
# date-time to string
#      "{0:yyyy-MM-dd} {0:HH:mm:ss}" -f (Get-Date)
#
# string to date-tim
#     [datetime]::parseexact('2020-10-10 23:45:00', 'yyyy-MM-dd HH:mm:ss', $null)
#     [datetime]::parse('2020-10-10 23:45:00', $null)
function minify_error_msg($err) {
    # reads an my custom error object (with .Err_MSG and .Err_ID properties)
    # and returns a rather short string describing the error
    $t = $err.Err_MSG
    # Con. to RPS0291-PC failed: WinRM cannot process the request. The following error with errorcode 0x80090322 occurred while using Kerberos authentication: An unknown security error occurred.    
    if ($err.Err_ID -eq '-2144108387,PSSessionStateBroken') {$t = 'Unknown kerberos security error 0x80090322'}
    $t = "$t" # in one line
    $t = $t -replace 'Connecting to remote server','Con. to'
    $t = $t -replace 'failed with the following error message :','failed:'
    $t = $t -replace 'The following error occurred','Err'
    if ($t.length -gt 200) {$t = $t.Substring(0,199)}
    $err.Err_ID + "; " + $t
}
function display_clients($list, $description, $color_if_not_empty, $color_if_empty, $dont_print_if_empty) {
    # just for logging it displays something like this:
    #        Clients {$description}: (3) cl1 cl2 cl3 
    if ($list) {
        $temp = list_clients_in_one_line $list
        log ("   Clients {0}: ({1}) $temp" -f $description, $list.count) $color_if_not_empty
    } else {
        if ($dont_print_if_empty) {} else {
        log ("   Clients {0}: (0)" -f $description) $color_if_empty}
    }
}       
function display_pending_clients() {
    # special case of display_clients for pending ones with fancy colloring
    #        Clients {$description}: (3) cl1 cl2 cl3 
    $ts = "{0:MM/dd} {0:HH:mm:ss}" -f (Get-Date)
    write-host "$ts $task_id " -NoNewline -ForegroundColor White
    if ($clients_pending) {
        Write-Host ("   Clients pending: ({0}) " -f $clients_pending.count) -NoNewline -ForegroundColor white
        foreach ($client in $clients_pending) {
            $fails = $failures_counts[$client]
            $client = $client -replace 'RPS','' -replace '-PC',''
            if ($fails) {
                if ($fails -ge $MAX_TASK_RETRIES) {
                    Write-Host "$client " -NoNewline -ForegroundColor yellow -BackgroundColor red
                } elseif ($fails -ge ($MAX_TASK_RETRIES/2)) {
                    Write-Host "$client " -NoNewline -ForegroundColor red
                } elseif ($fails -ge ($MAX_TASK_RETRIES/4)) {
                    Write-Host "$client " -NoNewline -ForegroundColor Magenta
                } else {
                    Write-Host "$client " -NoNewline -ForegroundColor yellow
                }
            } else {
                Write-Host "$client " -NoNewline -ForegroundColor DarkGray
            }
        }
        Write-Host ""
        $temp = list_clients_in_one_line $clients_pending
        log ("   Clients pending: ({0}) $temp" -f $clients_pending.count) white $true
    } else {
        if ($dont_print_if_empty) {} else {
        log "   Clients pending: 0 :-)" Green
        }
    }
} 
function report_of_pending_clients() {
    # a nice text report with one line per pending client like this:
    #  - RPS0323-PC    (4 failures)
    #  - RPS0325-PC    (never seen)
    $temp = @()
    ForEach ($client in $clients_pending) { `
        if ($failures_counts.Keys -contains $client) {
            $temp += (" - $client `t({0} failures)" -f $failures_counts[$client])
        } else {
            $temp += " - $client `t(never seen)"
        }
    }
    return $temp
}
function report_status_txt() {
    # returns a nice overal status report to rite to status.txt
    "Done clients:"
    list_clients_in_one_line $clients_done_alltime
    ""
    "Pending clients:"
    report_of_pending_clients
    ""
    "Started at:              {0}" -f (date_time_to_str $state['deployment_start'])
    "Last success at:         {0}" -f (date_time_to_str $state['last_success'])
    "---------CLIENTS---------------"
    "Online during last pass:  $len_clients_online"
    "Max online at once:      {0}" -f $state['max_clients_online_atonce']
    "Total:                    $len_clients_all "
    "Done:                     $len_clients_done_alltime"
    "Pending:                  $len_clients_pending = $len_clients_not_seen never seen + $len_failures failed"
}
function call_emar_for_all_tasks() {
    # for every task found in .\tasks\ execute 
    # $ emar run $task_id
    while (1) {
        Get-ChildItem "$base_dir\tasks\" -Directory -Exclude "_*" | ForEach-Object {
            Write-Host ""
            Write-Host "`"$PSScriptRoot\emar.ps1`" run $_.name" -ForegroundColor Cyan
            & "$PSScriptRoot\emar.ps1" run $_.name
        }
        # a quick'n'dirty count-down timer 
        18..1 | ForEach-Object {
            $sec=$_*10; Write-Host -NoNewLine "`rSleeping for ~$sec`"    `r"
            Start-Sleep 10}
        Write-Host -NoNewLine          "`r                     `r"
    }
}
function log_pass_and_overal_results() {
    # log status of this pass
    #----------------------------------------
    if ($clients_todo) {
        log ("In this pass (out of {0} clients todo)" -f $clients_todo.count)
        display_clients $clients_done_this_pass "done" Green Yellow
        display_clients $clients_failed "failed" yellow Gray
    }
    # log status of task (after all passes)
    #----------------------------------------
    log ("Since {0} (out of {1} clients)" -f (date_time_to_str $state['deployment_start']), $len_clients_all) 
    display_clients $clients_done_alltime "done"  DarkGray green      
    display_pending_clients  
    log "   Of the $len_clients_pending pending clients: $len_clients_not_seen have never been seen & $len_failures have failed" gray
    log ("   Last success was on {0}" -f (date_time_to_str $state['last_success'])) Gray
}

if (($command -eq "run") -and ($task_id)) {

    # INITIALISATION
    #-------------------------------------
    $task_dir = "$base_dir\tasks\$task_id"
    mkdir -force "$task_dir\results" > $null

    . "$task_dir\task.ps1"
    
    # Set default values for vars that are not set in task.ps1
    #---------------------------------------------------------
    if (!($Script:TIMEOUT)) {
        $default = 300
        log "Setting timeout to $default sec because task.ps1 doesnot set it, e.g. with: `$script:TIMEOUT = ..." yellow
        $Script:TIMEOUT = 300
    } 

    $state=@{}
    if (Test-Path -PathType Leaf "$task_dir\state_main.dat") {
        $state = load_state "$task_dir\state_main.dat"
    } else {
        log "Initializing task state because this is the first run" darkgray
        $state['deployment_start'] = (Get-Date)
        $state['last_success'] = $null
        $state['max_clients_online_atonce'] = 0   
        save_state $state "$task_dir\state_main.dat"
    }

    # $failures_counts is a dictionary with values like this:
    # "RPS1234-PC" : 5
    # Which means RPS1234-PC had 5 failed attempts since program start
    $failures_counts = @{}
    if (Test-Path -PathType Leaf "$task_dir\state_failures.dat") {
        #log "Loading failures count from $task_dir\state_failures.dat" darkgray
        $temp = (Get-Content "$task_dir\state_failures.dat" | ConvertFrom-Json)
        $temp.PSObject.Properties | ForEach-Object { $failures_counts[$_.Name]=$_.Value }
    }

    # init some vars
    # All clients_... vars are list of client names
    $major_client_errors = @() # those that we need to record in status.txt
    $clients_failed = @()
    $clients_done_this_pass = @()
    $clients_done_alltime = (Get-ChildItem "$task_dir\results\").Name `
        | ForEach-Object {$_ -replace '.txt',''}

    # load list of clients_all
    $clients_all = @(Get-Content $task_dir\clients.txt)

    log "Discovering clients" DarkGray
    $pings = (Test-Connection -ComputerName $clients_all -Count 1 -AsJob | Wait-Job | Receive-Job )
    get-job | remove-job   
    $pings_replying = $pings | Where-Object {$_.ResponseTime}
    $clients_online = $pings_replying | ForEach-Object {$_.Address}

    # FOR TESTING ONLU quick one-ofs):
    # Invoke-Command -ComputerName $clients_online -ScriptBlock {Get-PhysicalDisk | select 'FriendlyName','MediaType','BusType','Size','HealthStatus'}

    # clients "todo" are those that are online but have not succeeded the task
    $clients_todo = ($clients_online | Where-Object {$clients_done_alltime -notcontains $_})

    $len_clients_all=$clients_all.Count
    $len_clients_online=$clients_online.Count
    $len_clients_todo=$clients_todo.Count
    if ($len_clients_online -gt $state['max_clients_online_atonce']) {$state['max_clients_online_atonce'] = $len_clients_online}
    $len_clients_done_alltime=$clients_done_alltime.Count
    log "Clients: Total=$len_clients_all, Online=$len_clients_online (of which todo=$len_clients_todo)"
    display_clients $clients_online "online" "DarkGray" "yellow"
    display_clients $clients_todo "todo (before cleanup)" "DarkGray" "yellow" $true

    $cleanup_msg = "" # (for nice reporting only)
    if ($clients_todo) {
        # There are clients to-do (online clients that have not succeded the task)
        if ($failures_counts) {
            # Some clients have failures so we may skip some of the clients_todo
            # The possibility that we will try the task on one client decreases
            # as the failures of this client increase
            $clients_to_skip = @()
            foreach ($client in $clients_todo) {
                if ($failures_counts.Keys -contains $client) {
                    if ($failures_counts[$client] -ge 10) {
                        # I have >=10 failures for this client -- should I skip it?
                        $random = (Get-Random -Maximum $MAX_TASK_RETRIES)
                        $count = $failures_counts[$client]
                        if ($count -ge 50) {$count = 49}
                        if ($random -lt $count) {
                            # clients with 10 or more and up to $MAX_TASK_RETRIES 
                            # failures have decreasing
                            # probability of being handled a Task (80...2%)
                            # For 50 or more failures they have a 2% probability
                            $clients_to_skip += $client
            }}}}
            if ($clients_to_skip) {
                # we will not try clients in $clients_to_skip list
                $temp = (list_clients_in_one_line $clients_to_skip)
                log "Will skip these clients because they had too many failures: $temp" yellow
                $clients_todo = ($clients_todo | Where-Object { $clients_to_skip -notcontains $_})
                $cleanup_msg = " (after cleanup)"
    }}}
    $len_clients_todo=$clients_todo.Count
    display_clients $clients_todo "todo$cleanup_msg" "DarkGray" "Yellow"

    #----------------------------------------------
    # So, do we really have anything to do?
    #----------------------------------------------
    # $clients_todo = @("wx1-pc","bad1-pc","bad2-pc") # FOR TESTING ONLY
    # $clients_todo += "wx1-pc" # FOR TESTING ONLY 
    if ($clients_todo) {
        log ("Attempting to submit task to {0} clients" -f $clients_todo.Count) DarkGray
        Invoke-Command  -ThrottleLimit 30 -ComputerName $clients_todo -ScriptBlock ${Function:ClientTask} -AsJob  |
             Wait-Job -TimeOut $script:TIMEOUT  > $null
        $j = Get-Job
        $error.clear()
        $j | Receive-Job 2>$null >$null # this returns all the std-out of all jobs as one
        $results = $j.ChildJobs

        <# I collect all errors in a dictionary like this: 
            $client_err['client-name'] | fl
            Err_ID  : NetworkPathNotFound,PSSessionStateBroken
            Err_MSG : Connecting to remote server bad1-pc failed with ...
        
        BTW: This are the most interesting properties of $error:
            $error[0].TargetObject
                RPS0242-PC
            $error[0].FullyQualifiedErrorId
                -2144108387,PSSessionStateBroken
            $error[0].Exception.Message
                Connecting to remote server RPS0242-PC failed with the ... errorcode 0x80090322 occurred while ...
        #>
        #************************************************************
        # SENSITIVE PART OF CODE - AVOID MISTAKES UNDER THIS POINT
        #
        # I'm enumerating $error so any exception that happens here
        # will alter it as I enumerate it which is Not Good (TM)
        # - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        $client_err = @{}
        $other_err = @{}
        foreach ($err in $error) {
            $client  = $err.TargetObject
            if ($client) {
                $client_err[$client] = [PSCustomObject]@{
                    Err_ID     = $err.FullyQualifiedErrorId
                    Err_MSG    = $err.Exception.Message
                }
            } else {
                "Invoke-command error without reference to a client: {0}: {1}" `
                    -f $err.FullyQualifiedErrorId, $err.Exception.Message
                $other_err += "Invoke-command error without reference to a client: {0}: {1}" `
                    -f $err.FullyQualifiedErrorId, $err.Exception.Message
            }
        }
        # - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        # SENSITIVE PART OF CODE - AVOID MISTAKES UPTO THIS POINT
        #
        #************************************************************
        if ($other_err.Count) {$other_err | ForEach-Object {log $_ yellow}}

        if ($results) {
            # Nice we got results from Invoke-Command!
            foreach ($job in $results)  {
                $client=$job.Location
                $output = $job.output # FIXME maybe stderr is at $job.error - must check
                if (($job.state -eq 'Completed') -and ($output -match "<SUCCESS>$")) {
                    $output > "$task_dir\results\$client.txt"
                    $clients_done_this_pass += $client
                    $state['last_success'] = (Get-Date)
                    if (Test-Path "$task_dir\Bad.results.$client.txt") {Remove-Item "$task_dir\Bad.results.$client.txt"}
                } else {
                    # Either not completed (failed) or no <SUCCESS>
                    # First record failures
                    if ($client) {
                        if ($failures_counts.keys -contains $client) {
                            $failures_counts[$client] += 1
                        } else {
                            $failures_counts[$client] = 1
                        }
                    } else {
                        log "No `$job.Location in `$results (program logic error?)" red
                        $job | Format-List *
                    }
                    # Then log the error
                    $err = $client_err[$client]
                    if ($err) {
                        $failure_desc = minify_error_msg $err
                        # some errors are nothing to write home about
                        if ($err.Err_ID -match 'PSSessionStateBroken') {
                            $color = "white"
                            $prefix = "PSSession failure"
                        } else {
                            $major_client_errors += "$client $failure_desc"
                            $color = "red"
                            $prefix = "OTHER FAILURE"
                        }
                    } else {
                        $failure_desc="No PoSH exception but task didn't report <SUCCESS>. Output is:$output"
                        $color = "red"
                        $prefix = "MAJOR FAILURE"

                        $output > "$task_dir\Bad.results.$client.txt"
                        }
                    log "     $prefix, $client, $failure_desc" $color
                }
            } 
        } else {
            log "Got back NO RESULTS at all" yellow
        } # if ($results)

        # done parsing $results -- I can remove jobs
        get-job | remove-job
    } # if ($clients_todo)

    #-----------------------------------------------------------
    # Calculate lists, vars regarding  the status after this pass
    # (mainly for reporting)
    #-----------------------------------------------------------
    # update $failures_counts 
    # (remove any client that finally succeded in this pass)
    foreach ($client in $clients_done_this_pass) {
        if ($failures_counts.keys -contains $client) {
            $temp = $failures_counts[$client]
            if ($temp -gt 1) {log "NICE: Client $client completed the task after $temp failures"}
            $failures_counts.remove($client)
        }
    }
    foreach ($client in $client_err.Keys) {$clients_failed += $client}
    $clients_done_alltime = (Get-ChildItem "$task_dir\results\").Name | ForEach-Object {$_ -replace '.txt',''}
    $clients_pending = ($clients_all | Where-Object { $clients_done_alltime -notcontains $_ })
    # store the length of a few lists for easy reporting
    $len_clients_done_alltime = $clients_done_alltime.Count
    $len_clients_pending = $clients_pending.Count
    $len_failures = $failures_counts.Count
    $len_clients_not_seen = $len_clients_pending - $len_failures

    # log results
    #----------------------------------------
    log_pass_and_overal_results

    # Persit to disk (status, failures, state, pending)
    #----------------------------------------
    report_status_txt > "$task_dir\status.txt"
    $failures_counts | ConvertTo-Json > "$task_dir\state_failures.dat"
    save_state $state "$task_dir\state_main.dat"

    # end of one pass
    #----------------------------------------
} elseif ($command -eq "start") {
    call_emar_for_all_tasks
} else {
    Write-Host "Unknown command: $command" red
    call_emar_for_all_tasks
}