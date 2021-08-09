#!/bin/bash -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ROOT_DIR=$(cd ${SCRIPT_DIR}/..; pwd)

# create local env file if it doesn't exist

#[ ! -f "${ROOT_DIR}/.env" ] && cp "${ROOT_DIR}/.env.default" "${ROOT_DIR}/.env"

# source env to get platform specific docker compose command

. ${SCRIPT_DIR}/env.sh

# create containers

${DOCKER_COMPOSE_CMD} up -d --remove-orphans