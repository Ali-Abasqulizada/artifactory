#!/bin/sh

set -e
set -u

VERSION=${1} docker-compose  -H ssh://root@209.141.57.42  -f docker-compose.yml build artifactory_backend

VERSION=${1} docker-compose  -H ssh://root@209.141.57.42  -f docker-compose.yml up artifactory_backend   --renew-anon-volumes --remove-orphans -d
