requires 'App::Greple::msdoc';
requires 'App::Greple::subst';
requires 'App::sdif';
requires 'Data::Section::Simple';
requires 'Encode';
requires 'Getopt::EX::Long';
requires 'List::Util';
requires 'List::MoreUtils';
requires 'Moo';
requires 'Pod::Usage';
requires 'perl', 'v5.14.0';

on configure => sub {
    requires 'Module::Build::Tiny', '0.035';
};

on test => sub {
    requires 'Test::More', '0.98';
};
