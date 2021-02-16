DOCX  = $(wildcard *.docx)
NAME  = $(basename $(DOCX))

DIFF  = $(NAME:=.diff)
DIFF += $(NAME:=.small-diff)
STAT  = $(NAME:=.stat)

PATCH  = $(NAME:=.patch)
PATCH += $(NAME:=.pretty-patch)
PATCH += $(NAME:=.small-patch)
PATCH += $(NAME:=.small-pretty-patch)

ALL = $(DIFF) $(STAT) $(PATCH)

all: $(ALL)

$(ALL): Makefile

clean:
	rm -f $(ALL)

GREPLE = greple -Mmsdoc -Msubst --check=any
DIFF2VBA = perl -I../lib ../script/diff2vba

%.diff: %.docx
	$(GREPLE) --all-sample-dict --diff $< > $@

%.small-diff: %.docx
	$(GREPLE) --dictdata $$'コンピューター? コンピューター' --diff $< > $@


%.stat: %.docx
	$(GREPLE) --all-sample-dict --stat $< > $@


%.patch: %.diff
	$(DIFF2VBA) --no-pretty $< > $@

%.pretty-patch: %.diff
	$(DIFF2VBA) --pretty $< > $@

%.small-patch: %.small-diff
	$(DIFF2VBA) --no-pretty $< > $@

%.small-pretty-patch: %.small-diff
	$(DIFF2VBA) --pretty $< > $@

