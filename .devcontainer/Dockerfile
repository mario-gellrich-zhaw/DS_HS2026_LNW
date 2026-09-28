FROM mcr.microsoft.com/devcontainers/python:1-3.12-bookworm

# The Yarn apt repository in the base image has an invalid signing key, which
# breaks "apt-get update" and with it the docker-in-docker feature.
RUN grep -rl yarnpkg /etc/apt/sources.list.d/ | xargs -r rm -f
