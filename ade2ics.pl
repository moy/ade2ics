#!/usr/bin/env perl

# Author: Jean-Edouard BABIN, je.babin in telecom-bretagne.eu
# Extended by:
#	Ronan Keryell, rk in enstb.org
#	Matthieu Moy, Matthieu.Moy in grenoble-inp.fr
# 
# TODO
# - iCal output should respect standards more carefully
# - Base config should be in an external file
#
# For history see the end of the script

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

############################################
# Default school

# You may have to change this with your school name
# Script comes with configuration for TelecomBretagne and Ensimag (case sensitive)
# You may need to create a new configuration for you school bellow

my $default_school = 'TelecomBretagne';
############################################

# School configuration
my %default_config;
# For TelecomBretagne
$default_config{'TelecomBretagne'}{'u'} = 'http://edt.enst-bretagne.fr/ade/'; # (http://edt.enst-bretagne.fr/ade/ will not work)
$default_config{'TelecomBretagne'}{'l'} = '';
$default_config{'TelecomBretagne'}{'p'} = ''; # Should be commented if your ADE system don't need a password
$default_config{'TelecomBretagne'}{'t'} = 0;
$default_config{'TelecomBretagne'}{'d'} = 0;
$default_config{'TelecomBretagne'}{'s'} = 1;

# For Ensimag
$default_config{'Ensimag'}{'u'} = 'http://ade52-inpg.grenet.fr/ade/';
$default_config{'Ensimag'}{'l'} = 'voirIMATEL';
$default_config{'Ensimag'}{'p'} = ''; # Should be commented if your ADE system don't need a password
$default_config{'Ensimag'}{'t'} = 0;
$default_config{'Ensimag'}{'d'} = 0;
$default_config{'Ensimag'}{'s'} = 0;

my %opts;
my @tree;

# Output to UTF-8
binmode(STDOUT, ":encoding(UTF-8)");
$| = 1;

$opts{e} = $default_school;

if (!defined $ENV{REQUEST_METHOD}) {

	GetOptions(\%opts, 'c=s', 'u=s', 'l=s', 'e=s', 'p:s', 't!', 'd!', 's!');

	$opts{'u'} = $opts{'u'} || $default_config{$opts{'e'}}{'u'};
	$opts{'l'} = $opts{'l'} || $default_config{$opts{'e'}}{'l'};
	$opts{'p'} = $opts{'p'} || $default_config{$opts{'e'}}{'p'};
	$opts{'t'} = $opts{'t'} || $default_config{$opts{'e'}}{'t'};
	$opts{'d'} = $opts{'d'} || $default_config{$opts{'e'}}{'d'};
	$opts{'s'} = $opts{'s'} || $default_config{$opts{'e'}}{'s'};

	if (!defined($opts{'c'})) {
		print STDERR "Usage: $0 -c Path [-e school_name] [-u base_url] [-l login] [-p [password]] [-t] [-d] [-s]\n";
		print STDERR " -c is expecting the path through the page you need to click to get the information you are looking for, encoded in ISO-8859-1 (see examples)\n";
		print STDERR " -e is expecting you school name. It loads default value of -u -l -p -t- d- s for your school. Default school is : $default_school. Available school are (case sensitive):\n";
		print STDERR "\t- $_\n" foreach (keys %default_config);
		print STDERR " -u is expecting the ADE location to peek into\n";
		print STDERR " -l is expecting your login name for authentication purpose\n";
		print STDERR " -p is expecting the password to use for authentication purpose\n";
		print STDERR "\t 	if you just use -p without password, you will be prompted for it. recommanded for security !\n";
		print STDERR " -t write the schedule in time-stamped \"calendar.\" file to track modifications to your calendar.\n";
		print STDERR " -d enable verbose output\n";
		print STDERR " -s enable CAS Authentification, as used at Telecom Bretagne\n";
		print STDERR "\nSome examples:\n";
		print STDERR " $0 -l jebabin -p -c '2007-2008:Etudiants:FIP:FIP 3A 2007-2008:BABIN Jean-Edouard'\n";
		print STDERR " $0 -e Ensimag -p somepassword -c 'ENSIMAG2009-2010:Enseignants:M:Moy Matthieu'\n";
		print STDERR " even more:\n";
		print STDERR " $0 -s -l jebabin -p -c '2007-2008:Etudiants:FIP:FIP 3A 2007-2008:BABIN Jean-Edouard'\n";
		print STDERR " $0 -t -s -l keryell -p some_password -c '2007-2008:Enseignants:H Ã  K:KERYELL Ronan'\n";
		print STDERR " $0 -u http://ade52-inpg.grenet.fr/ade/ -l voirIMATEL -p somepassword -c 'ENSIMAG2009-2010:Enseignants:M:Moy Matthieu'\n";
		exit 1;
	}
} else {
	print header(-type => 'text/calendar; method=request; charset=UTF-8;', -attachment => 'edt.ics');
	if (defined(param('c'))) {
		$opts{'e'} = param('e') if (defined(param('e')));

		$opts{'u'} = $default_config{$opts{'e'}}{'u'};
		$opts{'l'} = $default_config{$opts{'e'}}{'l'};
		$opts{'p'} = $default_config{$opts{'e'}}{'p'};
		$opts{'t'} = $default_config{$opts{'e'}}{'t'};
		$opts{'d'} = $default_config{$opts{'e'}}{'d'};
		$opts{'s'} = $default_config{$opts{'e'}}{'s'};

		$opts{'c'} = param('c');
		$opts{'u'} = param('u') if (defined(param('u')));
		$opts{'l'} = param('l') if (defined(param('l')));
		$opts{'p'} = param('p') if (defined(param('p')));
		$opts{'t'} = param('t') if (defined(param('t')));
		$opts{'s'} = param('s') if (defined(param('s')));
		$opts{'d'} = param('d') if (defined(param('d')));
	} else {
		print "Usage: $0?c=Chemin[&e=school][&u=base_url][&l=login][&p=password][&t][&d]\n";
		exit 1;
	}
}

if ((defined($opts{'p'})) and ($opts{'p'} eq "")) {
	print "Please input password: ";
	ReadMode('noecho');
	$opts{'p'} = ReadLine(0);
	chomp $opts{'p'};
	ReadMode('normal');
}

if ($opts{'t'}) {
    # Create a time stamped output file:
    my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime();
    my $output_file_name = sprintf("calendar.%d-%02d-%02d_%02d:%02d:%02d.ics",
				   $year+1900, $mon, $mday, $hour, $min, $sec);
	close(STDOUT);
    open(STDOUT, ">", $output_file_name) or die "open of " . $output_file_name . " failed!";
}

foreach (split(':', $opts{'c'})) {
	push (@tree,$_);
	# $opts{'c'} is ProjectID:Category:branchId:branchId:branchId but of course in text instead of ID
	# The job is to find the latest branchId ID then parse schedule
}

my $mech = WWW::Mechanize->new(agent => 'ADEics 0.2', cookie_jar => {});


# login in
$mech->get($opts{'u'}.'standard/index.jsp');
die "Error 1 : check if base_url work" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{'d'});

if ($opts{'s'}) {
	$mech->submit_form(fields => {username => $opts{'l'}, password => $opts{'p'}});
} else {
	$mech->submit_form(fields => {login => $opts{'l'}, password => $opts{'p'}});
}
print STDERR $mech->content."\n" if ($opts{'d'});
die "Error 2" if (!$mech->success());

if ($opts{'s'}) {
	$mech->follow_link( n => 1 );
	print STDERR $mech->content."\n" if ($opts{'d'});
	die "Error 2.1" if (!$mech->success());
}

# Getting projet list
$mech->get($opts{'u'}.'standard/projects.jsp');
die "Error 2.2 : check if ADE url ($opts{'u'}) works" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{'d'});

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
print STDERR $mech->content."\n" if ($opts{'d'});


# We need to load tree.jsp to find category name
$mech->get($opts{'u'}.'standard/gui/tree.jsp');
die "Error 5" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{'d'});

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
$mech->get($opts{'u'}.'standard/gui/tree.jsp?category='.$category.'&expand=false&forceLoad=false&reload=false&scroll=0');
die "Error 7" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{'d'});


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
	print STDERR $mech->content."\n" if ($opts{'d'});
	die "Error 8.$_ : $tree[$_] does not exist" if (!defined($branchId));

	if ($_ == $#tree) {
		$mech->get($opts{'u'}.'standard/gui/tree.jsp?selectId='.$branchId.'&reset=true&forceLoad=false&scroll=0');
	} else {
		$mech->get($opts{'u'}.'standard/gui/tree.jsp?branchId='.$branchId.'&expand=false&forceLoad=false&reload=false&scroll=0');
	}
}

print STDERR $mech->content."\n" if ($opts{'d'});
die "Error 9 : $tree[$#tree] does not exist" if (!defined($branchId));

# We need to choose a week
$mech->get($opts{'u'}.'custom/modules/plannings/pianoWeeks.jsp?forceLoad=true');
die "Error 10" if (!$mech->success());
print $mech->content."\n" if ($opts{'d'});

# then we choose all week
$mech->get($opts{'u'}.'custom/modules/plannings/pianoWeeks.jsp?searchWeeks=all');
die "Error 10bis" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{'d'});

# Get planning
$mech->get($opts{'u'}.'custom/modules/plannings/info.jsp');
die "Error 11" if (!$mech->success());
print STDERR $mech->content."\n" if ($opts{'d'});

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

	while ($token = $p->get_tag("tr")) {
		my $date;
		my $id;
		my $course;
		my $week;
		my $day;
		my $hour;
		my $duration;
		my $trainers;
		my $trainees;
		my $rooms;
		my $equipment;
		my $statuts;
		my $groupes;
		my $module;
		my $formation_UV;

		#######################################
		# This part is not generic enough to work well with all installation
		#######################################

		$token = $p->get_tag("span");
		$date = $p->get_trimmed_text; # 12/05/2006

		$token = $p->get_tag("a");
		$id = $token->[1]{href};
		$id =~ /\((\d+)\)/;
		$id = $1;

		$course = $p->get_trimmed_text; # INF 423 Cours 1 et 2
		$token = $p->get_tag("td");
		$week = $p->get_trimmed_text; # 10 sept. 2007 | S40-09
		$token = $p->get_tag("td");
		$day = $p->get_trimmed_text; # Mardi
		$token = $p->get_tag("td");
		$hour = $p->get_trimmed_text; # 13h30 | 15:30
		$token = $p->get_tag("td");
		$duration = $p->get_trimmed_text; # 2h50min | 2h | 50min

		$token = $p->get_tag("td");
		$trainees = $p->get_text('td'); #
		$token = $p->get_tag("td");
		$trainers = $p->get_text('td'); # LEROUX Camille
		$token = $p->get_tag("td");
		$rooms = $p->get_text('td'); # B03-132A
		$token = $p->get_tag("td");
		$equipment = $p->get_text('td');
		$token = $p->get_tag("td");
		$statuts = $p->get_text('td'); # Validé
		$token = $p->get_tag("td");
		$groupes = $p->get_text('td'); # Groupe UV2 MAJ INF 423
		$token = $p->get_tag("td");
		$module = $p->get_text('td'); # FIP ELP103 Electronique num?rique : Logique combinatoire
		$token = $p->get_tag("td");
		$formation_UV = $p->get_text('td'); # Enseignements INF S3 UV2 MAJ INF Automne Majeure INF UV2
		
		#######################################
		if(0) { #used for debug
		print "Date:		$date\n";
		print "Id:		$id\n";
		print "course:		$course\n";
		print "week:		$week\n";
		print "day:		$day\n";
		print "hour:		$hour\n";
		print "duration:	$duration\n";
		print "trainers:	$trainers\n";
		print "trainees:	$trainees\n";
		print "rooms:		$rooms\n";
		print "equipment:	$equipment\n";
		print "statuts:		$statuts\n";
		print "groupes:	$groupes\n";
		print "module:		$module\n";
		print "formation_UV:	$formation_UV\n";
		print "\n";
		next;
		}
		#######################################
		
		$date =~ m|(\d+)/(\d+)/(\d+)|;
		my $ics_day = sprintf("%02d%02d%02d",$3,$2,$1);
		$hour =~ m|(\d+)[h:](\d+)|;
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
		print "Cours : $course".'\n' if ($course !~ /^\s+$/);
		print "Etudiants : $trainees".'\n' if ($trainees !~ /^\s+$/);
		print "Groupes	: $groupes".'\n' if ($groupes !~ /^\s+$/);
		print "Modules	: $module".'\n' if ($module !~ /^\s+$/);
		print "Formations/UV : $formation_UV".'\n' if ($formation_UV !~ /^\s+$/);
		print "Equipements : $equipment".'\n' if ($equipment !~ /^\s+$/);
		print "Statuts	: $statuts\n" if ($statuts !~ /^\s+$/);
		print "LOCATION:$rooms\n";
		print "URL;VALUE=URI:".$opts{'u'}."custom/modules/plannings/eventInfo.jsp?eventId=$id\n";
		print "END:VEVENT\n";
	}
}

__END__

History (doesn't follow commit revision)

Revision 2.7 2009/09/03
Allow -s -d -t switch to be negated with -nos --nod --not
Add an example for Ensimag
Add -e switch to select base configuration with a school name.
Minor bugfix

Revision 2.6 2008/09/19
Handle date with : instead of h

Revision 2.6 2008/09/10
Bug fix ($opts{'p'})

Revision 2.5 2008/09/09
Password can be read from stdin

Revision 2.4 2008/09/09
In-line help message improvment (keryell)

Revision 2.3 2007/09/18
Should now work with outlook 2002 (Thanks to C. Lohr)

Revision 2.2 2007/09/09
Can now be used with CAS authentification
Updated to work with new EDT installation

Revision 2.1 2007/03/14
Now send 'edt.ics' filename when using HTTP

Revision 2.0 2007/03/06
New way to get planning (much much faster) with info.jsp
TZID change for client that don't understand Europe/Paris

Revision 1.5 2006/10/11
Simpler way to get all weeks (thanks Erka !)
Change debug message to STDERR

Revision 1.4 2006/10/11
Show all weeks, including current one. bad code, but works...
Change OUTPUT to STDOUT, closing it then opening it to a file.

Revision 1.3 2006/09/26
Keep all the records in the calendar output.
Cleaned up output routine.
Rename the output file with an .ics extension.

Revision 1.2 2006/09/26
Improved documentation.
Now try to catch all the weeks in the calendar (TODO: to query first the
existing weeks or use the all weeks option of EDT?...).
Added a time-stamped output file option with -t.
