[![Actions Status](https://github.com/kaz-utashiro/App-diff2vba/workflows/test/badge.svg)](https://github.com/kaz-utashiro/App-diff2vba/actions)
# NAME

diff2vba - generate VBA patch script from diff output

# VERSION

Version 0.06

# SYNOPSIS

greple -Mmsdoc -Msubst \\
    --all-sample-dict --diff some.docx | diff2vba > patch.vba

# DESCRIPTION

**diff2vba** is a command to generate VBA patch script from diff output.

# OPTIONS

- **--maxlen**=_n_

    Set maximum length of literal string.
    Default is 250.

- **--subname**=_name_

    Set subroutine name in the VBA script.
    Default is `Patch`.

- **--identical**

    Produce patch script for identical string.
    Default is false.

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
