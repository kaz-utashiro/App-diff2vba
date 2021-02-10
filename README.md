[![Build Status](https://travis-ci.com/kaz-utashiro/App-diff2vba.svg?branch=master)](https://travis-ci.com/kaz-utashiro/App-diff2vba)
# NAME

diff2vba - generate VBA patch script from diff output

# VERSION

Version 0.01

# SYNOPSIS

greple -Msubst --diff old.docx new.docx | diff2vba > patch.vba

# DESCRIPTION

**diff2vba** is a command to generate VBA patch script from diff output.

# OPTIONS

- **--fold**\[=_width_\]

    Fold VBA string literals.  Default true with 72 column width.
    Specify 0 not to fold.

- **--variable**\[=_name_\]

    Specify VBA array variable name.  Default is `subst`.

# INSTALL

cpanm https://github.com/kaz-utashiro/App-diff2vba.git

# SEE ALSO

[App::Greple](https://metacpan.org/pod/App::Greple), [https://github.com/kaz-utashiro/greple](https://github.com/kaz-utashiro/greple)

[App::Greple::subst](https://metacpan.org/pod/App::Greple::subst), [https://github.com/kaz-utashiro/greple-subst](https://github.com/kaz-utashiro/greple-subst)

[https://qiita.com/kaz-utashiro/items/85add653a71a7e01c415](https://qiita.com/kaz-utashiro/items/85add653a71a7e01c415)

# AUTHOR

Kazumasa Utashiro

# LICENSE

Copyright 2021 Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.
