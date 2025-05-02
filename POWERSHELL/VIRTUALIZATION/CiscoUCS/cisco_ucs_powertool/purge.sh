#!/bin/bash

set -ex

# docker hub username
USERNAME=slmingol
# image name
IMAGE=cisco_ucs_powertool

docker images $USERNAME/$IMAGE -q | sort | xargs docker rmi -f
