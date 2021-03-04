package App::diff2vba;
use 5.014;
use warnings;

our $VERSION = "0.04";

use utf8;
use Encode;
use Data::Dumper;
{
    no warnings 'redefine';
    *Data::Dumper::qquote = sub { qq["${\(shift)}"] };
    $Data::Dumper::Useperl = 1;
}
use open IO => 'utf8', ':std';
use Pod::Usage;
use Data::Section::Simple qw(get_data_section);
use List::Util qw(max);
use List::MoreUtils qw(pairwise);
use App::diff2vba::Util;
use App::sdif::Util;

use Moo;

has debug     => ( is => 'ro' );
has verbose   => ( is => 'ro', default => 1 );
has format    => ( is => 'ro', default => 'dumb' );
has subname   => ( is => 'ro', default => 'Patch' );
has maxlen    => ( is => 'ro', default => 250 );
has identical => ( is => 'ro', default => undef );
has quotes    => ( is => 'rw',
		   default => sub { { '“' => "Chr(&H8167)", '”' => "Chr(&H8168)" } } );
has quotes_re => ( is => 'rw' );

no Moo;

sub run {
    my $app = shift;
    local @ARGV = map { utf8::is_utf8($_) ? $_ : decode('utf8', $_) } @_;

    use Getopt::EX::Long qw(GetOptions Configure ExConfigure);
    ExConfigure BASECLASS => [ __PACKAGE__, "Getopt::EX" ];
    Configure qw(bundling no_getopt_compat);
    GetOptions($app, make_options "
	debug
	verbose | v !
	format      =s
	subname     =s
	maxlen      =i
	identical   !
	") || pod2usage();

    $app->initialize;

    for my $file (@ARGV ? @ARGV : '-') {
	$app->process_file($file);
    }
}

sub substitute {
    my $app = shift;
    my $script = sprintf "subst_%s.vba", $app->format;
    my $max = $app->maxlen;
    for my $fromto (@_) {
	use integer;
	chomp @$fromto;
	my($s_from, $s_to) = @$fromto;
	my $longer = max(map length, $s_from, $s_to);
	my $count = ($longer + $max - 1) / $max;
	my @from = _split_string($s_from, $count);
	my @to   = _split_string($s_to,   $count);
	for my $i (0 .. $#from) {
	    next if !$app->identical and $from[$i] eq $to[$i];
	    print $app->section($script,
				{ TARGET      => $app->vba_string_literal($from[$i]),
				  REPLACEMENT => $app->vba_string_literal($to[$i]) });
	}
    }
}

sub process_file {
    my $app = shift;
    my $file = shift;

    open my $fh, $file or die "$file: $!\n";

    my @fromto;
    while (<$fh>) {
	#
	# diff --combined (generic)
	#
	if (m{^
	       (?<command>
	       (?<mark> \@{2,} ) [ ]
	       (?<lines> (?: [-+]\d+(?:,\d+)? [ ] ){2,} )
	       \g{mark}
	       (?s:.*)
	       )
	       }x) {
	    my($command, $lines) = @+{qw(command lines)};
	    my $column = length $+{mark};
	    my @lines = map {
		$_ eq ' ' ? 1 : int $_
	    } $lines =~ /\d+(?|,(\d+)|( ))/g;

	    warn $_ if $app->{debug};

	    next if @lines != $column;
	    next if $column != 2;

	    push @fromto, $app->read_diff($fh, @lines);
	}
    }

    printf "Sub %s()\n\n", $app->subname;
    print $app->section("setup.vba");
    $app->substitute(@fromto);
    print "End Sub\n";
}

sub initialize {
    my $app = shift;
    my $chrs = join '', keys %{$app->quotes};
    $app->quotes_re(qr/[\Q$chrs\E]/);
}

sub section {
    my $app = shift;
    my $section = shift;
    my $replace = shift // {};
    local $_ = get_data_section($section);
    for my $name (keys %$replace) {
	s/\b(\Q$name\E)\b/$replace->{$1}/ge;
    }
    $_;
}

sub read_diff {
    my $app = shift;
    my($fh, @lines) = @_;
    my @diff = read_unified_2 $fh, @lines;
    my @out;
    while (my($c, $o, $n) = splice(@diff, 0, 3)) {
	@$o > 0 and @$o == @$n or next;
	s/^[\t +-]// for @$c, @$o, @$n;
	push @out, pairwise { [ $a, $b ] } @$o, @$n;
    }
    @out;
}

sub vba_string_literal {
    my $app = shift;
    my $quotes = $app->quotes;
    my $chrs_re = $app->quotes_re;
    join(' & ',
	 map { $quotes->{$_} || sprintf('"%s"', s/\"/\"\"/gr) }
	 map { split /($chrs_re)/ } @_);
}

######################################################################

sub _split_string {
    local $_ = shift;
    my $count = shift;
    my $len = int((length($_) + $count - 1) / $count);
    my @split;
    while (length) {
	push @split, substr($_, 0, $len, '');
    }
    @split == $count or die;
    @split;
}

1;

=encoding utf-8

=head1 NAME

App::diff2vba - generate VBA patch script from diff output

=head1 SYNOPSIS

greple -Msubst --diff old.docx new.docx | diff2vba > patch.vba

=head1 DESCRIPTION

B<diff2vba> is a command to generate VBA patch script from diff output.

=head1 OPTIONS

=over 7

=item B<--maxlen>=I<n>

Set maximum length of literal string.
Default is 250.

=begin comment

=item B<--format>=I<format>

Set format of VBA script.
Default is C<dumb>.

=end comment

=item B<--subname>=I<name>

Set subroutine name in the VBA script.
Default is C<Patch>.

=item B<--identical>

Produce patch script for identical string.
Default is false.

=back

=head1 AUTHOR

Kazumasa Utashiro

=head1 LICENSE

Copyright 2021 Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

__DATA__

@@ setup.vba

With Selection.Find
    .MatchCase = True
    .MatchByte = True
    .IgnoreSpace = False
    .IgnorePunct = False
End With

@@ subst_dumb.vba

With Selection.Find
    .Text = TARGET
    .Replacement.Text = REPLACEMENT
    .Execute Replace:=wdReplaceOne
End With
Selection.Collapse Direction:=wdCollapseEnd

@@ subst_dumb2.vba

With Selection.Find
    .Text = TARGET
    if .Execute Then
        Selection.Range.Text = REPLACEMENT
    End If
End With

