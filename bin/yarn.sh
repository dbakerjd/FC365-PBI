#!/bin/bash -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ROOT_DIR=$(cd ${SCRIPT_DIR}/..; pwd)

# source env to get platform specific docker compose command

. ${SCRIPT_DIR}/env.sh

CONTAINER=${1}
shift

CMD="yarn $@"

${DOCKER_COMPOSE_CMD} exec ${CONTAINER} sh -c "${CMD}"

