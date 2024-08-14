#!/bin/bash

set -ex

# docker hub username
USERNAME=slmingol
# image name
IMAGE=cisco_ucs_powertool

docker build -t $USERNAME/$IMAGE:latest .
