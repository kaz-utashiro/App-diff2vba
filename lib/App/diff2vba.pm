package App::diff2vba;
use 5.014;
use warnings;

our $VERSION = "0.01";



1;
__END__

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

