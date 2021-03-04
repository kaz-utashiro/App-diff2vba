[![Actions Status](https://github.com/kaz-utashiro/App-diff2vba/workflows/test/badge.svg)](https://github.com/kaz-utashiro/App-diff2vba/actions)
# NAME

diff2vba - generate VBA patch script from diff output

# VERSION

Version 0.04

# SYNOPSIS

greple -Mmsdoc -Msubst \\
    --all-sample-dict --diff some.docx | diff2vba > patch.vba

# DESCRIPTION

**diff2vba** is a command to generate VBA patch script from diff output.

# OPTIONS

- **--fold**\[=_width_\]

    Fold string literals.  Default is 72.  Specify 0 not to fold.

- **--variable**=_name_

    Specify array variable name.  Default is `subst`.

- **--**\[**no-**\]**pretty**

    Default true and produce script with newlines and indentation for
    human readability.  If disabled, produce concise data without them.

# INSTALL

cpanm https://github.com/kaz-utashiro/App-diff2vba.git

# SEE ALSO

[App::Greple](https://metacpan.org/pod/App::Greple), [https://github.com/kaz-utashiro/greple](https://github.com/kaz-utashiro/greple)

[App::Greple::msdoc](https://metacpan.org/pod/App::Greple::msdoc), [https://github.com/kaz-utashiro/greple-msdoc](https://github.com/kaz-utashiro/greple-msdoc)

[App::Greple::subst](https://metacpan.org/pod/App::Greple::subst), [https://github.com/kaz-utashiro/greple-subst](https://github.com/kaz-utashiro/greple-subst)

[https://qiita.com/kaz-utashiro/items/85add653a71a7e01c415](https://qiita.com/kaz-utashiro/items/85add653a71a7e01c415)

# AUTHOR

Kazumasa Utashiro

# LICENSE

Copyright 2021 Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.
