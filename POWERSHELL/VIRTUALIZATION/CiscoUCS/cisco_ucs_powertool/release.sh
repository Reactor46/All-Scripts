#!/bin/bash

set -ex

# docker hub username
USERNAME=slmingol
# image name
IMAGE=cisco_ucs_powertool

# ensure we're up to date
#git pull

version=`cat VERSION`
echo "version: $version"

# run build
./build.sh

# tag it
#git add -A
#git commit -m "version $version"
#git tag -a "$version" -m "version $version"
#git push
#git push --tags

docker tag $USERNAME/$IMAGE:latest $USERNAME/$IMAGE:$version

# push it
docker push $USERNAME/$IMAGE:latest
docker push $USERNAME/$IMAGE:$version

