#!/usr/bin/env zsh -f
# Purpose: Look for company name in Microsoft Word Documents and tag file with name
#
# From:	Timothy J. Luoma
# Mail:	luomat at gmail dot com
# Date:	2019-12-29

NAME="$0:t:r"

if [[ -e "$HOME/.path" ]]
then
	source "$HOME/.path"
else
	PATH="$HOME/scripts:/usr/local/bin:/usr/bin:/usr/sbin:/sbin:/bin"
fi

zmodload zsh/datetime

TIME=$(strftime "%Y-%m-%d-at-%H.%M.%S" "$EPOCHSECONDS")

LOG="$HOME/Library/Logs/$NAME.log"

[[ -d "$LOG:h" ]] || mkdir -p "$LOG:h"
[[ -e "$LOG" ]]  || touch "$LOG"

function timestamp { strftime "%Y-%m-%d-at-%H:%M:%S" "$EPOCHSECONDS" }

function log { echo "$NAME [`timestamp`]: $@" | tee -a "$LOG" }

if ((! $+commands[tag] ))
then
		# note: if tag is a function or alias, it will come back not found
	log "'tag' is required but not found in $PATH"
	log "You can install 'tag' with 'brew install tag' or from <https://github.com/jdberry/tag/>"

	exit 0
fi

for i in "$@"
do

	[[ ! -e "$i" ]] && log "'$i' does not exist." && continue

	[[ ! -f "$i" ]] && log "'$i' is not a file." && continue

	[[ ! -w "$i" ]] && log "'$i' is not writable." && continue

	EXT="$i:e:l"

	[[ "$EXT" == "" ]] && log "'$i' has no file extension." && continue

	case "$EXT" in
		'doc'|'docx'|'xls'|'ppt')
			# Put all of the extensions that you want to try to act on in the list above
			:
		;;

		*)
			log "'$i' has an unrecognized file extension '$EXT'."
			continue
		;;
	esac

	# if we get here then we have one of the 'good' file extensions

		# this should get us the company name, if any
	COMPANY_NAME=$(mdls "$i" | awk "/kMDItemOrganizations/{getline; print}" | awk -F'"' '//{print $2}')

	[[ "$COMPANY_NAME" == "" ]] && log "'$i' has no company name." && continue

		# check to see if the tag already exists
	tag --garrulous --no-name --tags "$i" \
		| egrep -q "^${COMPANY_NAME}$" \
		&& log "'$i' already has the tag '$COMPANY_NAME'." \
		&& continue

	# if we get here, we need to tag the file

	tag --add "$COMPANY_NAME" "$i"

	EXIT="$?"

	if [[ "$EXIT" == "0" ]]
	then
		log "'$i' was successfully tagged '$COMPANY_NAME'."
	else
		log "FAILED to tag '$i' as '$COMPANY_NAME'."
	fi

done

exit 0
#EOF
