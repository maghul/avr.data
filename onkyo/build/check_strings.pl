#!/bin/perl -w


my $stringPrefix= $ARGV[0];

print "Checking '$stringPrefix'\n";

sub add {
	my ( $map, $string, $ref )= @_;

	my $refs = (exists $$map{$string})?$$map{$string}:[];
	push @$refs, $ref;
	$$map{$string}= $refs;
}

sub loadFromFile {
	my ( $map, $filename ) = @_;

	open INFILE, $filename;
	my $ii=0;
	while (<INFILE>) {
		++$ii;
		chomp;
		if ( /$stringPrefix/ ) {
			my @strings= split( /$stringPrefix/ );
			shift @strings;
			for my $string (  @strings ) {
				$string =~ s/[^A-Z0-9_].*//;
				$string = $stringPrefix.$string;
				add( $map, $string, "$filename:$ii" );
			}
		}
	}
}



sub loadUsedStrings {
	my $map= {};
	
	open FIND, "find . | ";
	while ( <FIND> ) {
		chomp;
		if ( /.htm$/ || /.html$/ || /.pm$/ || /.pl$/ || /install.xml/ ) {
			loadFromFile( $map, $_ );
		}
	}
	return $map
}

sub loadDefinedStrings {
	my $map= {};
	
	loadFromFile( $map, "strings.txt" );

	return $map
}

sub common {
	my ( $a, $b ) = @_;
	my $elms= {};
	for my $elma ( keys %$a ) {
		$$elms{$elma}=1 if ( $$b{$elma} );
	}
	return $elms;
}

sub eliminate {
	my ( $a, $b ) = @_;
	for my $elma ( keys %$a ) {
		delete $$b{$elma};
	}
}

sub report {
	my ( $map, $missing ) = @_;
	my $flag;

	for my $elm ( sort keys %$map ) {
		my $list= $$map{$elm};
		for my $ref ( @$list ) {
			if ( $missing ) {
				print "$ref: Missing string '$elm'\n";
				$flag= 1;
			} else {
				print "$ref: Unused string '$elm'\n";
			}
		}
	}
	exit if $flag;
}

my $used_strings= loadUsedStrings();

my $defined_strings= loadDefinedStrings();

my $common_strings= common( $used_strings, $defined_strings );

eliminate( $common_strings, $used_strings );
eliminate( $common_strings, $defined_strings );

report( $used_strings, 1 );
report( $defined_strings, 0 );
