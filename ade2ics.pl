#!/usr/bin/env perl -w

# Author: Jean-Edouard BABIN, je.babin in telecom-bretagne.eu
# Extended by Ronan Keryell, rk in enstb.org
# 
# Redistribution of this script, is prohibited without the
# agreement of the author.
# 
# TODO
# - iCal output should respect standards more carefully

# $Log: edt.pl,v $
# Revision 2.5 2008/09/09 22:30:48 jeb
# Password can be read from stdin
#
# Revision 2.4 2008/09/09 21:35:12 jeb
# In-line help message improvment (keryell)
#
# Revision 2.3 2007/09/18 17:12:24 jeb
# Should now work with outlook 2002 (Thanks to C. Lohr)
#
# Revision 2.2 2007/09/09 14:49:32 jeb
# Can now be used with CAS authentification
# Updated to work with new EDT installation
# 
# Revision 2.1 2007/03/14 01:54:46 jeb
# Now send 'edt.ics' filename when using HTTP
#
# Revision 2.0 2007/03/06 22:55:33 jeb
# New way to get planning (much much faster) with info.jsp
# TZID change for client that don't understand Europe/Paris
#
# Revision 1.5 2006/10/11 22:18:42 jeb
# Simpler way to get all weeks (thanks Erka !)
# Change debug message to STDERR
#
# Revision 1.4 2006/10/11 18:11:28 jeb
# Show all weeks, including current one. bad code, but works...
# Change OUTPUT to STDOUT, closing it then opening it to a file.
#
# Revision 1.3 2006/09/26 12:31:42  keryell
# Keep all the records in the calendar output.
# Cleaned up output routine.
# Rename the output file with an .ics extension.
#
# Revision 1.2 2006/09/26 08:53:30 keryell
# Improved documentation.
# Now try to catch all the weeks in the calendar (TODO: to query first the
# existing weeks or use the all weeks option of EDT?...).
# Added a time-stamped output file option with -t.

# Revision 1.1 2006/09/19 15:43:28 keryell
# Initial revision

use strict;
use warnings;

use WWW::Mechanize;
use HTTP::Cookies;
use HTML::TokeParser;
use Term::ReadKey;
use Getopt::Long;
use Time::Local;
use POSIX;

use CGI qw(param header);

my %opts;
my @tree;

# Output to UTF-8
binmode(STDOUT, ":encoding(UTF-8)");
$| = 1;

# Base config
my $default_url = $opts{u} = 'http://edt.enst-bretagne.fr/ade/'; # Don't work with edt.telecom-bretagne.eu !
$opts{l} = '';
$opts{p} = '';
$opts{t} = 0;
$opts{d} = 0;
$opts{s} = 1;

if (!defined $ENV{REQUEST_METHOD}) {

	GetOptions(\%opts, 'c=s', 'u=s', 'l=s', 'p:s', 't', 'd', 's');

	if (!defined($opts{c})) {
		print STDERR "Usage: $0 -c Chemin [-u base_url] [-l login] [-p [password]] [-t] [-d] [-s]\n";
		print STDERR " Chemin is the path through the page you need to click to get the information you are looking for, encoded in ISO-8859-1\n";
		print STDERR " -d for verbose output\n";
		print STDERR " -t to write the schedule in time-stamped \"calendar.\" file to track modifications to your calendar.\n";
		print STDERR " -u base_url : the ADE location to peek into. Default value is \"-u $default_url\"\n";
		print STDERR " -s to use CAS Authentification, as used at Telecom Bretagne\n";
		print STDERR " -l login : define your login name for authentication purpose (at Telecom Bretagne you need to put your account name)\n";
		print STDERR " -p password : define the password to use for authentication purpose (at Telecom Bretagne you need to put your password name)\n";
		print STDERR "\t 	if you just use -p without password, you will be prompted for it. recommanded for security !\n";
		print STDERR "\nSome examples:\n $0 -s -l jebabin -p -c '2007-2008:Etudiants:FIP:FIP 3A 2007-2008:BABIN Jean-Edouard'\n";
		print STDERR " $0 -t -s -l keryell -p some_password -c '2007-2008:Enseignants:H à K:KERYELL Ronan'\n";
		exit 1;
	}
} else {
	print header(-type => 'text/calendar; method=request; charset=UTF-8;', -attachment => 'edt.ics');
	$opts{u} = param('u') if (defined(param('u')));
	$opts{l} = param('l') if (defined(param('l')));
	$opts{p} = param('p') if (defined(param('p')));
	$opts{t} = param('t') if (defined(param('t')));
	$opts{s} = param('s') if (defined(param('s')));
	$opts{d} = param('d') if (defined(param('d')));
	if (defined(param('c'))) {
		$opts{c} = param('c');
	} else {
		print "Usage: $0?c=Chemin&[u=base_url]&[l=login]&[p=password]&[t]&[d]\n";
	}
}

if ((defined($opts{p})) and ($opts{p} eq "")) {
	print "Please input password: ";
	ReadMode('noecho');
	$opts{p} = ReadLine(0);
	chomp $opts{p};
	ReadMode('normal');
}

if ($opts{t}) {
    # Create a time stamped output file:
    my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime();
    my $output_file_name = sprintf("calendar.%d-%02d-%02d_%02d:%02d:%02d.ics",
				   $year+1900, $mon, $mday, $hour, $min, $sec);
	close(STDOUT);
    open(STDOUT, ">", $output_file_name) or die "open of " . $output_file_name . " failed!";
}

foreach (split(':', $opts{c})) {
	push (@tree,$_);
	# $opts{c} is ProjectID:Category:branchId:branchId:branchId but of course in text instead of ID
	# The job is to find the latest branchId ID then parse schedule
}

my $mech = WWW::Mechanize->new(agent => 'ADEics 0.2', cookie_jar => {});


# login in
$mech->get($opts{u}.'standard/index.jsp');
die "Error 1 : check if base_url work" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{d});

if ($opts{s}) {
	$mech->submit_form(fields => {username => $opts{l}, password => $opts{p}});
} else {
	$mech->submit_form(fields => {login => $opts{l}, password => $opts{p}});
}
print STDERR $mech->content."\n" if ($opts{d});
die "Error 2" if (!$mech->success());

if ($opts{s}) {
	$mech->follow_link( n => 1 );
	print STDERR $mech->content."\n" if ($opts{d});
	die "Error 2.1" if (!$mech->success());
}

# Getting projet list
$mech->get($opts{u}.'standard/projects.jsp');
die "Error 2.2 : check if base_url work" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{d});

# Choosing projectId
my $p = HTML::TokeParser->new(\$mech->content);
my $token = $p->get_tag("select");

my $projid = -1;
while (($projid == -1) && (my $token = $p->get_tag("option"))) {
      if($p->get_trimmed_text eq $tree[0]) {
	      $projid = $token->[1]{value};
	}
}
die "Error 3 : $tree[0] does not exist" if ($projid == -1);

$mech->submit_form(fields => {projectId => $projid});
die "Error 4" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{d});


# We need to load tree.jsp to find category name
$mech->get($opts{u}.'standard/gui/tree.jsp');
die "Error 5" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{d});

# So, finding it
$p = HTML::TokeParser->new(\$mech->content);
$token = $p->get_tag("div");

my $category;
while ((!defined($category)) && (my $token = $p->get_tag("a"))) {
      if($p->get_trimmed_text eq $tree[1]) {
	      $category = $token->[1]{href};
	}
}
$category =~ s/.*\('(.*?)'\)$/$1/;
die "Error 6 : $tree[1] does not exist" if (!defined($category));


# We need load the category chosed on command line to find branchID
$mech->get($opts{u}.'standard/gui/tree.jsp?category='.$category.'&expand=false&forceLoad=false&reload=false&scroll=0');
die "Error 7" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{d});


# We loop until last branchID
my $branchId;
for (2..$#tree) {
	undef $branchId;

	# find branch
	$p = HTML::TokeParser->new(\$mech->content);
	$token = $p->get_tag("div");

	while ((!defined($branchId)) && (my $token = $p->get_tag("a"))) {
		if($p->get_trimmed_text eq $tree[$_]) {
		      $branchId = $token->[1]{href};
		}
	}
	$branchId =~ s/.*\((\d+),\s+.*/$1/;
	print STDERR $mech->content."\n" if ($opts{d});
	die "Error 8.$_ : $tree[$_] does not exist" if (!defined($branchId));

	if ($_ == $#tree) {
		$mech->get($opts{u}.'standard/gui/tree.jsp?selectId='.$branchId.'&reset=true&forceLoad=false&scroll=0');
	} else {
		$mech->get($opts{u}.'standard/gui/tree.jsp?branchId='.$branchId.'&expand=false&forceLoad=false&reload=false&scroll=0');
	}
}

print STDERR $mech->content."\n" if ($opts{d});
die "Error 9 : $tree[$#tree] does not exist" if (!defined($branchId));

# We need to choose a week
$mech->get($opts{u}.'custom/modules/plannings/pianoWeeks.jsp?forceLoad=true');
die "Error 10" if (!$mech->success());
print $mech->content."\n" if ($opts{d});

# then we choose all week
$mech->get($opts{u}.'custom/modules/plannings/pianoWeeks.jsp?searchWeeks=all');
die "Error 10bis" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{d});

# Get planning
$mech->get($opts{u}.'custom/modules/plannings/info.jsp');
die "Error 11" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{d});

# Parse planning to get event
$p = HTML::TokeParser->new(\$mech->content);

print "BEGIN:VCALENDAR\n";
print "VERSION:2.0\n";
print "PRODID:-//Jeb//edt.pl//EN\n";
print "BEGIN:VTIMEZONE\n";
#print "TZID:Europe/Paris\n";
#print "X-WR-TIMEZONE:Europe/Paris\n";
print "TZID:\"GMT +0100 (Standard) / GMT +0200 (Daylight)\"\n";
print "END:VTIMEZONE\n";
print "METHOD:PUBLISH\n";
ics_output($mech->content, $1);
print "END:VCALENDAR\n";

sub ics_output {
    # Parse the data and generate records that tends toward
    # http://www.ietf.org/rfc/rfc2445.txt :-)
	my $data = $_[0];
	my $p = HTML::TokeParser->new(\$data);
	
	$token = $p->get_tag("table");
	$token = $p->get_tag("tr");
	$token = $p->get_tag("tr");

	my $i = 0;
	while ($token = $p->get_tag("tr")) {
		$token = $p->get_tag("span");
		my $date = $p->get_trimmed_text; # 12/05/2006

		$token = $p->get_tag("a");
		my $id = $token->[1]{href};
		$id =~ /\((\d+)\)/;
		$id = $1;

		my $course = $p->get_trimmed_text; # INF 423 Cours 1 et 2
		$token = $p->get_tag("td");
		my $week = $p->get_trimmed_text; # 10 sept. 2007
		$token = $p->get_tag("td");
		my $day = $p->get_trimmed_text; # Mardi
		$token = $p->get_tag("td");
		my $hour = $p->get_trimmed_text; # 13h30
		$token = $p->get_tag("td");
		my $duration = $p->get_trimmed_text; # 2h50min | 2h | 50min

		$token = $p->get_tag("td");
		my $type = $p->get_trimmed_text; # C | TP | ...
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		$p->get_trimmed_text; # ??
		$token = $p->get_tag("td");
		my $trainees = $p->get_trimmed_text; #
		$token = $p->get_tag("td");
		my $trainers = $p->get_trimmed_text; # LEROUX Camille
		$token = $p->get_tag("td");
		my $rooms = $p->get_text('td'); # B03-132A
		$token = $p->get_tag("td");
		my $equipment = $p->get_text('td');
		$token = $p->get_tag("td");
		my $statuts = $p->get_text('td'); # Valid�
		$token = $p->get_tag("td");
		my $groupes = $p->get_text('td'); # Groupe UV2 MAJ INF 423
		$token = $p->get_tag("td");
		my $module = $p->get_trimmed_text; # FIP ELP103 Electronique num?rique : Logique combinatoire
		$token = $p->get_tag("td");
		my $formation_UV = $p->get_trimmed_text; # Enseignements INF S3 UV2 MAJ INF Automne Majeure INF UV2

		$date =~ m|(\d+)/(\d+)/(\d+)|;
		my $ics_day = sprintf("%02d%02d%02d",$3,$2,$1);
		$hour =~ m|(\d+)h(\d+)|;
		my $ics_hour .= sprintf("%02d%02d00",$1,$2);
		my $ics_start_date = $ics_day.'T'.$ics_hour;
	
		my $ics_stop_date;
		if ($duration =~ m|^(\d+)h(\d+)|) {
			$ics_stop_date = $ics_day.'T'.sprintf('%06d',($ics_hour+($1*10000)+($2*100))); # Ok this is wrong as we can get minute > 59, but ical understand it, that's ok for now
		} elsif ($duration =~ m|^(\d+)h|) {
			$ics_stop_date = $ics_day.'T'.sprintf('%06d',($ics_hour+($1*10000)));
		} elsif ($duration =~ m|^(\d+)m|) {
			$ics_stop_date = $ics_day.'T'.sprintf('%06d',($ics_hour+($1*100)));
		} else {
			die "Error 14 : date $duration can't be parsed";
		}
	
		my ($tssec,$tsmin,$tshour,$tsmday,$tsmon,$tsyear,$tswday,$tsyday,$tsisdst) = gmtime();
                my $dtstamp = sprintf("%02d%02d%02dT%02d%02d%02dZ", $tsyear+1900, $tsmon + 1, $tsmday, $tshour, $tsmin, $tssec);

		print "BEGIN:VEVENT\n";
		print "DTSTART:$ics_start_date\n";
		print "DTEND:$ics_stop_date\n";
		print "SUMMARY:$course\n";
		print "DTSTAMP:$dtstamp\n";
		print "UID:edt-$id-0\n";		
		print "DESCRIPTION:";
		print "Salle : $rooms".'\n';
		print "Enseignants : $trainers".'\n';
		print "Cours : $course".'\n' if ($course ne '-');
		print "Etudiants : $trainees".'\n' if ($trainees ne '-');
		print "Groupes	: $groupes".'\n';
		print "Modules	: $module".'\n';
		print "Formations/UV : $formation_UV".'\n';
		print "Equipements : $equipment".'\n' if ($equipment ne '-');
		print "Statuts	: $statuts\n";
		print "LOCATION:$rooms\n";
		print "URL;VALUE=URI:".$opts{u}."custom/modules/plannings/eventInfo.jsp?eventId=$id\n";
		print "END:VEVENT\n";
		$i++;
	}
}
