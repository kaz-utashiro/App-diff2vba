package App::diff2vba;
use 5.014;
use warnings;

our $VERSION = "0.03";

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
use Text::ANSI::Fold qw(:constants);
use Text::VisualPrintf qw(vprintf vsprintf);
use Data::Section::Simple qw(get_data_section);
use List::MoreUtils qw(pairwise);
use App::diff2vba::Util;
use App::sdif::Util;

use Moo;

has debug     => ( is => 'ro' );
has verbose   => ( is => 'ro', default => 1 );
has format    => ( is => 'ro', default => 'default' );
has pretty    => ( is => 'ro', default => 2 );
has fold      => ( is => 'rw', default => 72 );
has boundary  => ( is => 'ro', default => 'word' );
has linebreak => ( is => 'ro', default => LINEBREAK_ALL );
has margin    => ( is => 'ro', default => 4 );
has indent    => ( is => 'rw', default => 2 );
has variable  => ( is => 'ro', default => "subst" );
has connect   => ( is => 'rw', default => "\n" );

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
	pretty      !
	fold        :72
	margin      =i
	indent      =i
	variable    =s
	connect     =s
	") || pod2usage();

    $app->initialize;

    for my $file (@ARGV ? @ARGV : '-') {
	$app->process_file($file);
    }
}

my %driver = (
    default => sub {
	my $app = shift;
	my @fromto = @_;
	my $dim = 
	    sprintf("Dim %s (,) As String = ", $app->variable) .
	    $app->produce_array(@fromto);
	print _continue($dim), "\n";
	print $app->section('subst_one.vba', { VAR => $app->variable } );
    },
    dumb => sub {
	my $app = shift;
	my @fromto = @_;
	for my $fromto (@fromto) {
	    my($from, $to) = @$fromto;
	    print $app->section('subst_dumb.vba',
				{ TARGET      => _continue($app->vba_string_literal($from)),
				  REPLACEMENT => _continue($app->vba_string_literal($to)) });
	}
    },
    );

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

	    push @fromto, $app->process_diff(read_unified_2 $fh, @lines);
	}
    }

    print "Sub Patch()\n\n";
    $driver{$app->format || 'default'}->($app, @fromto);
    print "End Sub\n";
}

sub initialize {
    my $app = shift;
    if (not $app->pretty) {
	$app->fold(0);
	$app->indent(0);
	$app->connect('');
    }
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

sub process_diff {
    my $app = shift;
    my @out;
    while (my($c, $o, $n) = splice(@_, 0, 3)) {
	@$o > 0 and @$o == @$n or next;
	s/^[\t +-]// for @$c, @$o, @$n;
	push @out, pairwise { [ $a, $b ] } @$o, @$n;
    }
    @out;
}

sub produce_array {
    my $app = shift;
    my @pairs = map { $app->subst_pairs(@$_) } @_;
    $app->enclose_list(@pairs);
}

sub subst_pairs {
    my $app = shift;
    my($o, $n) = @_;
    my $s;
    my $indent = ' ' x $app->indent;
    $app->enclose_list($app->vba_string_literal($o),
		       $app->vba_string_literal($n));
}

sub vba_string_literal {
    my $app = shift;
    chomp(my $s = shift);
    return sprintf qq'"%s"', _vba_string($s) unless $app->fold;
    state $fold = $app->fold_obj;
    my @s = do {
	map { qq'"$_"' }
	map { _vba_string($_) }
	$fold->text($s)->chops;
    };
    my $width = $app->fold + $app->margin + length('""');
    join("\n",
	 map({ vsprintf "%-*s &", $width, $_ } splice @s, 0, -1),
	 @s);
}

sub fold_obj {
    my $app = shift;
    Text::ANSI::Fold->new(width     => $app->fold,
			  boundary  => $app->boundary,
			  linebreak => $app->linebreak,
			  runin     => $app->margin,
			  runout    => $app->margin);
}

sub enclose_list {
    my $app = shift;
    my $c = $app->connect;
    join($c, "{", $app->indent_text(join($c, _join_list(",", @_))), "}");
}

sub indent_text {
    my $app = shift;
    my $indent = ' ' x $app->indent;
    $_[0] =~ s/^/$indent/mgr;
}

######################################################################

sub _vba_string {
    local $_ = shift;
    s/"/""/g;
    $_;
}

sub _join_list {
    my $by = shift;
    return () if @_ < 1;
    (shift, map { ( $by, $_ ) } @_);
}

sub _continue {
    $_[0] =~ s/\n/ _\n/gr;
}

1;

=encoding utf-8

=head1 NAME

App::diff2vba - generate VBA patch script from diff output

=head1 SYNOPSIS

    greple -Msubst --diff old.docx new.docx | diff2vba > patch.vba

=head1 DESCRIPTION

B<diff2vba> is a command to generate VBA patch script from diff output.

=head1 AUTHOR

Kazumasa Utashiro

=head1 LICENSE

Copyright 2021 Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

__DATA__

@@ subst_one.vba

For index = 0 To VAR.GetUpperBound(0)
    With Selection.Find
        .Text = VAR(index, 0)
        .Replacement.Text = VAR(index, 1)
        .Execute Replace:=wdReplaceOne
    End With
Next

@@ subst_all.vba

For index = 0 To VAR.GetUpperBound(0)
    With Selection.Find
        .Text = VAR(index, 0)
        Do While .Execute
            Selection.Range.Text = VAR(index, 1)
        Loop
    End With    
Next

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

