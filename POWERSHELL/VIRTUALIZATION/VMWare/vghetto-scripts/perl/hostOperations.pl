#!/usr/bin/perl -w
# Author: William Lam
# Website: www.virtuallyghetto.com
# Reference: http://communities.vmware.com/docs/DOC-11435

use strict;
use warnings;
use VMware::VIRuntime;
use VMware::VILib;

# define custom options for vm and target host
my %opts = (
   'hostfile' => {
      type => "=s",
      help => "List of hosts to perform operation on",
      required => 1,
   },
   'operation' => {
      type => "=s",
      help => "ESX(i) Operation to perform [ent_maint|ext_maint|reboot]",
      required => 1,
   },
);

# read and validate command-line parameters 
Opts::add_options(%opts);
Opts::parse();
Opts::validate();

# connect to the server and login
Util::connect();

my ($host_view,$task_ref,$hostfile);
my @host_list = ();
$hostfile = Opts::get_option("hostfile");

&processFile($hostfile);

if( Opts::get_option('operation') eq 'ent_maint' ) {
	foreach my $host_name( @host_list ) {
                chomp($host_name);

		$host_view = Vim::find_entity_view(view_type => 'HostSystem', filter => { 'name' => $host_name});

		eval {
			$task_ref = $host_view->EnterMaintenanceMode_Task(timeout => 0, evacuatePoweredOffVms => 'true');
			print "Entering maintenance mode for host: \"" . $host_view->name . "\" and evacauating any VMs if host is part of DRS Cluster ...\n";
			my $msg = "\tSuccessfully entered maintenance mode for host: \"" . $host_view->name . "\"!";
			&getStatus($task_ref,$msg);
		};
		if ($@) {
        		# unexpected error
			print "Error: " . $@ . "\n\n";
		}
	}
} elsif(Opts::get_option('operation') eq 'ext_maint' ) {
	foreach my $host_name( @host_list ) {
                chomp($host_name);

                $host_view = Vim::find_entity_view(view_type => 'HostSystem', filter => { 'name' => $host_name});

                eval {
                        $task_ref = $host_view->ExitMaintenanceMode_Task(timeout => 0);
                        print "Exiting maintenance mode for host: \"" . $host_view->name . "\" ...\n";
                        my $msg = "\tSuccessfully exited maintenance mode for host: \"" . $host_view->name . "\"!";
                        &getStatus($task_ref,$msg);
                };
                if ($@) {
                        # unexpected error
                        print "Error: " . $@ . "\n\n";
                }
        }
} elsif(Opts::get_option('operation') eq 'reboot' ) {
	foreach my $host_name( @host_list ) {
                chomp($host_name);

                $host_view = Vim::find_entity_view(view_type => 'HostSystem', filter => { 'name' => $host_name});

		if($host_view->runtime->inMaintenanceMode) {
	                eval {
                	        $task_ref = $host_view->RebootHost_Task(force => 0);
        	                print "Rebooting host: \"" . $host_view->name . "\" ...\n";
	                        my $msg = "\tSuccessfully rebooted host: \"" . $host_view->name . "\"!";
                        	&getStatus($task_ref,$msg);
                	};
        	        if ($@) {
	                        # unexpected error
                        	print "Error: " . $@ . "\n\n";
                	}
		} else {
			print "Error: Host " . $host_view->name . " is not in maintenance mode!\n\n";
		}
        }
}

Util::disconnect();

sub getStatus {
	my ($taskRef,$message) = @_;
	
	my $task_view = Vim::get_view(mo_ref => $taskRef);
	my $taskinfo = $task_view->info->state->val;
	my $continue = 1;
	while ($continue) {
        	my $info = $task_view->info;
        	if ($info->state->val eq 'success') {
			print $message,"\n";
               		$continue = 0;
        	} elsif ($info->state->val eq 'error') {
               		my $soap_fault = SoapFault->new;
	   	        $soap_fault->name($info->error->fault);
               		$soap_fault->detail($info->error->fault);
               		$soap_fault->fault_string($info->error->localizedMessage);
               		die "$soap_fault\n";
        	}
        	sleep 5;
        	$task_view->ViewBase::update_view_data();
	}
}

# Subroutine to process the input file
sub processFile {
        my ($vmlist) =  @_;
        my $HANDLE;
        open (HANDLE, $vmlist) or die(timeStamp('MDYHMS'), "ERROR: Can not locate \"$vmlist\" input file!\n");
        my @lines = <HANDLE>;
        my @errorArray;
        my $line_no = 0;

        close(HANDLE);
        foreach my $line (@lines) {
                $line_no++;
                &TrimSpaces($line);

                if($line) {
                        if($line =~ /^\s*:|:\s*$/){
                                print "Error in Parsing File at line: $line_no\n";
                                print "Continuing to the next line\n";
                                next;
                        }
                        my $host = $line;
                        &TrimSpaces($host);
                        push @host_list,$host;
                }
        }
}

sub TrimSpaces {
        foreach (@_) {
                s/^\s+|\s*$//g
        }
}
