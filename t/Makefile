DOCX  = $(filter-out ~%,$(wildcard *.docx))
NAME  = $(basename $(DOCX))

DIFF  = $(NAME:=.diff)
DIFF += $(NAME:=.small-diff)
STAT  = $(NAME:=.stat)

PATCH  = $(NAME:=.patch)
PATCH += $(NAME:=.small-patch)
PATCH += $(NAME:=.dumb-patch)
PATCH += $(NAME:=.pretty-patch)
PATCH += $(NAME:=.small-pretty-patch)
PATCH += $(NAME:=.small-dumb-patch)

SJIS = $(PATCH:=.sjis)

ALL = $(DIFF) $(STAT) $(PATCH)

all: $(ALL)

sjis: $(SJIS)

$(ALL): Makefile

clean:
	rm -f $(ALL) $(SJIS)

GREPLE = greple -Mmsdoc -Msubst --check=any
DIFF2VBA = perl -I../lib ../script/diff2vba

%.diff: %.docx
	$(GREPLE) --all-sample-dict --diff $< | sed 1,2d > $@

%.small-diff: %.docx
	$(GREPLE) --dict small.dict --diff $< | sed 1,2d > $@


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

%.dumb-patch: %.diff
	$(DIFF2VBA) --no-pretty --format=dumb $< > $@

%.small-dumb-patch: %.small-diff
	$(DIFF2VBA) --no-pretty --format=dumb $< > $@

%patch.sjis: %patch
	iconv -f utf8 -t sjis < $< > $@
