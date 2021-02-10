requires 'App::Greple::msdoc';
requires 'App::Greple::subst';
requires 'App::sdif';
requires 'Data::Section::Simple';
requires 'Encode';
requires 'Getopt::EX::Long';
requires 'Moo';
requires 'Pod::Usage';
requires 'Text::ANSI::Fold';
requires 'Text::VisualPrintf';
requires 'perl', 'v5.14.0';

on configure => sub {
    requires 'Module::Build::Tiny', '0.035';
};

on test => sub {
    requires 'Test::More', '0.98';
};
