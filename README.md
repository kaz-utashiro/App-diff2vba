[![Actions Status](https://github.com/kaz-utashiro/App-diff2vba/workflows/test/badge.svg)](https://github.com/kaz-utashiro/App-diff2vba/actions)
# NAME

diff2vba - generate VBA patch script from diff output

# SYNOPSIS

    optex -Mtextconv diff -u A.docx B.docx | diff2vba > patch.vba

    greple -Mmsdoc -Msubst \
        --all-sample-dict --diff some.docx | diff2vba > patch.vba

# VERSION

Version 0.09

# DESCRIPTION

**diff2vba** is a command to generate VBA patch script from diff output.

# OPTIONS

- **--maxlen**=_n_

    Because VBA script does not accept string longer than 255 characters,
    longer string have to be splitted into shorter ones.  This option
    specifies maximum length.  Default is 250.

- **--adjust**=_n_

    Adjust border when the splitted strings are slightly different.
    Default value is 2, so they are adjusted upto two characters.  Set
    zero to disable it.

- **--subname**=_name_

    Set subroutine name in the VBA script.
    Default is `Patch`.

- **--identical**

    Produce patch script for identical string.
    Default is false.

- **--reverse**

    Generate reverse order patch.

# INSTALL

cpanm https://github.com/kaz-utashiro/App-diff2vba.git

# SEE ALSO

[App::Greple](https://metacpan.org/pod/App::Greple), [https://github.com/kaz-utashiro/greple](https://github.com/kaz-utashiro/greple)

[App::Greple::msdoc](https://metacpan.org/pod/App::Greple::msdoc), [https://github.com/kaz-utashiro/greple-msdoc](https://github.com/kaz-utashiro/greple-msdoc)

[App::Greple::subst](https://metacpan.org/pod/App::Greple::subst), [https://github.com/kaz-utashiro/greple-subst](https://github.com/kaz-utashiro/greple-subst)

[App::optex::textconv](https://metacpan.org/pod/App::optex::textconv), [https://github.com/kaz-utashiro/optex-textconv](https://github.com/kaz-utashiro/optex-textconv)

[https://qiita.com/kaz-utashiro/items/06c60843213b0f024df7](https://qiita.com/kaz-utashiro/items/06c60843213b0f024df7),
[https://qiita.com/kaz-utashiro/items/85add653a71a7e01c415](https://qiita.com/kaz-utashiro/items/85add653a71a7e01c415)

# AUTHOR

Kazumasa Utashiro

# LICENSE

Copyright 2021 Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.
