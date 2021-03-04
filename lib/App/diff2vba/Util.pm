package App::diff2vba;
use v5.14;
use warnings;

sub make_options {
    map {
	# "foo_bar" -> "foo_bar|foo-bar|foobar"
	s{^(?=\w+_)(\w+)\K}{
	    "|" . $1 =~ tr[_][-]r . "|" . $1 =~ tr[_][]dr
	}er;
    }
    grep {
	s/#.*//;
	s/\s+//g;
	/\S/;
    }
    map { split /\n+/ }
    @_;
}

sub split_string {
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
