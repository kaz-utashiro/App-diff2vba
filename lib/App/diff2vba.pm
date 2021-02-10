package App::diff2vba;
use 5.014;
use warnings;

our $VERSION = "0.01";

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
use App::diff2vba::Util;
use App::sdif::Util;

use Moo;

has debug     => ( is => 'ro' );
has verbose   => ( is => 'ro', default => 1 );
has fold      => ( is => 'ro', default => 72 );
has boundary  => ( is => 'ro', default => 'word' );
has linebreak => ( is => 'ro', default => LINEBREAK_ALL );
has margin    => ( is => 'ro', default => 4 );
has shift     => ( is => 'ro', default => 2 );
has variable  => ( is => 'ro', default => "subst" );

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
	fold        :72
	margin      =i
	shift       =i
	variable    =s
	") || pod2usage();

    $app->initialize;

    for my $file (@ARGV ? @ARGV : '-') {
	$app->process_file($file);
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

	    push @fromto,
		map { $app->process_diff($_) } read_unified $fh, @lines;
	}
    }

    my $data = 
	sprintf("Dim %s (,) As String = ", $app->variable) .
	$app->produce_script(@fromto);
    print $data =~ s/\n(?!\z)/ _\n/mgr;

    print $app->section('subst_one.vba');
}

sub initialize {
    my $app = shift;
}

sub section {
    my $app = shift;
    local $_ = get_data_section(+shift);
    s/\{\{\s*(.*?)\s*\}\}/$1/gee;
    $_;
}

sub process_diff {
    my $app = shift;
    my($buf, $mark_re) = @_;
    my @c = $buf->collect(qr/^[\t ]+$/);
    my @o = $buf->collect(qr/[-]/);
    my @n = $buf->collect(qr/[+]/);
    @o > 0 and @n > 0 and @o == @n or return;
    s/^[\t +-]// for @c, @o, @n;
    map { [ $o[$_], $n[$_] ] } 0 .. $#o;
}

sub produce_script {
    my $app = shift;
    my @pairs = map { $app->pairs(@$_) } @_;
    my $script =
	"{\n" .
	$app->indent(join(",\n", @pairs)) . "\n" .
	"}\n";
}

sub pairs {
    my $app = shift;
    my($o, $n) = @_;
    my $s;
    my $indent = ' ' x $app->shift;
    join("\n",
	 "{",
	 $app->indent($app->vba_string_literal($o) . "\n,\n" .
		      $app->vba_string_literal($n)),
	 "}");
}

sub indent {
    my $app = shift;
    my $indent = ' ' x $app->shift;
    $_[0] =~ s/^/$indent/mgr;
}

sub vba_string_literal {
    my $app = shift;
    chomp(my $s = shift);
    return sprintf qq'"%s"', vba_string($s) unless $app->fold;
    state $fold = $app->fold_obj;
    my @s = do {
	map { qq'"$_"' }
	map { _vba_string($_) }
	$fold->text($s)->chops;
    };
    my $width = $app->fold + $app->margin + 2;
    join '',
	map({ vsprintf "%-*s &\n", $width, $_ } splice @s, 0, -1),
	@s;
}

sub fold_obj {
    my $app = shift;
    Text::ANSI::Fold->new(width     => $app->fold,
			  boundary  => $app->boundary,
			  linebreak => $app->linebreak,
			  runin     => $app->margin,
			  runout    => $app->margin);
}

######################################################################

sub _vba_string {
    local $_ = shift;
    s/"/""/g;
    $_;
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

For index = 0 To {{ $app->variable }}.GetUpperBound(0)
    With Selection.Find
        .Text = {{ $app->variable }}(index, 0)
        .Replacement.Text = {{ $app->variable }}(index, 1)
        .Execute Replace:=wdReplaceOne
    End With
Next

@@ subst_all.vba

For index = 0 To {{ $app->variable }}.GetUpperBound(0)
    With Selection.Find
        .Text = {{ $app->variable }}(index, 0)
        Do While .Execute
            Selection.Range.Text = {{ $app->variable }}(index, 1)
        Loop
    End With    
Next

