
use strict;
use Spreadsheet::XLSX;
use Data::Dumper;
use Text::Iconv;

#my $FileName = "ISCP\ AV\ Receiver\ v124-1.xls";
#my $FileName = "ISCP-V1.26_2013.xls";
#my $FileName = "ISCP-V1.26_patched_2013.xls";
my $FileName = "ISCP_AVR_133 (430940).xlsx";
my $converter = Text::Iconv -> new ("utf-8", "windows-1251");
my $workbook   = Spreadsheet::XLSX->new($FileName, $converter);

my $source= $ARGV[0];
my $file= $ARGV[1];

#die $parser->error(), ".\n" if ( !defined $workbook );

# Following block is used to Iterate through all worksheets
# in the workbook and print the worksheet content 

my $content= {};

# Really sloooow
sub eliminateDuplicateStrings {
	my $str= shift;

	my $substr= substr($str, 0, 12);
	my @subs = split(quotemeta($substr), $str);

	my $count= @subs;
	my $tail= $subs[1];
#	print "count=$count\n";
#	print "head=$substr\n";
#	print "tail=$tail\n";
#	print "count=$count\n";
	if ($count>2) {
			my $res= join("", $substr,$tail);
			print "res=$res\n";
			return $res;
	}
	return $str;
}

#my $source= "Hejsan Hoppsan FalleralleraHejsan Hoppsan FalleralleraHejsan Hoppsan Fallerallera";
#my $res= eliminateDuplicateStrings($source);
#print "res='$res'\n";
#exit 0;

for my $worksheet ( $workbook->worksheets() ) {
    
    my $wsname= $worksheet->get_name();
    print "Worksheet --- $wsname\n";

    if ( $wsname=~/CMND/ ) {
	# Find out the worksheet ranges
	my ( $row_min, $row_max ) = $worksheet->row_range();
	my ( $col_min, $col_max ) = $worksheet->col_range();
	
	my $devices= [];
	for my $col ( $col_min+2 .. $col_max ) {
	    my $cell = $worksheet->get_cell( $row_min+1, $col );
	    last unless $cell;

	    my $device= $cell->value();
	    $device=~ s/TX-/ TX-/g;
	    $device=~ s/D TX-/ DTX-/g;
	    $device=~ s/DTR-/ DTR-/g;
	    $device=~ s/DHC-/ DHC-/g;
	    $device=~ s/\s\([Ee]ther\)/ /g;
	    $device=~ s/\s+/ /g;
	    my $prefix= $device;
	    $prefix=~ s/[0-9].*//;
	    $device=~ s/ \// $prefix/;
	    $device=~ s/  *\(/(/g;

	    # Typos...
	    $device=~ s/TX-NR5000ETX-NA1000/TX-NR5000E TX-NA1000/g;

	    if ( $device ) {
#		print "DEVICE $col --> '$device'\n";
		$$devices[$col]= $device;
	    } else {
		    last;
	    }
	}

	my $base;
	my $base_descr;
	my $base_cmd;
	for my $row ( $row_min+2 .. $row_max ) {
	    my $cell= $worksheet->get_cell( $row, 0 );
	    my $cmd= $cell?$cell->value():"None";
#	    print "cmd: $cmd\n";
	    if ( $cmd=~/\s*\"[A-Z0-9]{3,3}\"\s*-\s*/ ) {
		$cmd =~ s/^\s*//;
		$cmd =~ s/\s*$//;
		print "Base Command '$cmd'\n";
		$base= $cmd;
		$base_cmd= $cmd;
		$base_descr= $base;
		$base_descr =~ s/[^\s]* - //;
		
		$base_cmd =~ s/^"//;
		$base_cmd =~ s/".*//;

	    } else {
		$cmd =~ s/["\x{201c}\x{201d}]//g;  # Strip quotes
		$cmd =~ s/\x{2026}/.../g;
		$cmd =~ s/\s*//g;
		my $cell= $worksheet->get_cell( $row, 1 );
		my $descr= $cell?$cell->value():"None";
		$descr =~ s/\x{2013}/.../g;
		$descr =~ s/\x{2026}/.../g;
		$descr =~ s/\x{2019}/'/g;
		$descr =~ s/&gt;/>/msg;
		$descr =~ s/&lt;/</msg;
		$descr =~ s/&amp;/&/msg;
		$descr =~ s/\r\n/\n/sg;
		$descr =~ s/["\x{201c}\x{201d}]//g;  # Strip quotes
		$descr =~ s/\205/.../msg;
		$descr =~ s/\223(.*)\224/$1/msg;
		$descr =~ s/\223([^\224]*)\224/$1/msg;
		
		$cmd =~ s/&gt;/>/msg;
		$cmd =~ s/&lt;/</msg;
		$cmd =~ s/&amp;/&/msg;
		$cmd =~ s/\205/.../msg;

		print "   NitWit Command '$cmd' --> '$descr'\n";
		$cmd =~ s/\223([^\224]*)\224/$1/msg;
		
		# Nitwit elimination...
#		$descr =~ s/ChannelInfomation.*/Channel/ms;
#		$descr =~ s/Modeinfomation.*/Mode/ms;
#		$descr =~ s/maxset Keyboard.*/max/ms;
#		$descr =~ s/URLwaiting.*/URL/ms;
#		$descr =~ s/max\)NET\/USB.*/max)/ms;
		$descr = eliminateDuplicateStrings($descr);
		
		# Embarrasing...
		$descr =~ s/infomation/information/ms;
		$descr =~ s/Infomation/Information/ms;
		$cmd =~ s/infomation/information/ms;
		$cmd =~ s/Infomation/Information/ms;

		print "   UnWit Command '$cmd' --> '$descr'\n";
		for my $col ( $col_min+2 .. $col_max ) {
		    
		    # Return the cell object at $row and $col
		    my $cell = $worksheet->get_cell( $row, $col );
		    next unless $cell;
		    
#		    print "($row, $col) --> '".$cell->value()."'\n";

		    if ( $cell->value() =~ /[Yy]es/ ) {
			my $devlist= $$devices[$col];
			my $prefix= "";    
#			print "RR DEVICES: $devlist\n";
			for my $dev (split(/  */, $devlist)) {
				my $tmpprefix= $dev;
				$tmpprefix=~ s/[0-9].*//;
				if ($tmpprefix eq "/") {
					$dev =~ s/\//$prefix/
				} else {
					$prefix= $tmpprefix;
				}
#				print "RR DEVICE: '$dev'\n";
				my $devcont= $$content{$dev} || {};
				my $basecont= $$devcont{$base} || { "Description" => $base_descr, "Command" => $base_cmd, "Arguments" => {} };
				my $args= $$basecont{"Arguments"};
				$$args{$descr}=$cmd;
				$$devcont{$base}=$basecont;
				$$content{$dev}=$devcont;
			}
		    }
		}
	    }
	}
    }
}


for my $device ( keys %$content ) {
#	print "ZZ DEVICE: '$device'\n";
	open OUTPUT, "> ../$device";
	my $d = Data::Dumper->new( [$$content{$device}], ["\$\$device{'$device'}"] );
	$d->Sortkeys(1);
	$d->Indent(1);
	print OUTPUT $d->Dump();
#	print OUTPUT Dumper( $$content{$device} );
	close OUTPUT;
}
