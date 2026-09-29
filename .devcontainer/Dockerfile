# Pinned by digest so the environment cannot change before the assessment.
FROM mcr.microsoft.com/devcontainers/python:1-3.12-bookworm@sha256:7876580dc67fd460fd962f004cbeb480027e9bbc0657096f1087db11f9eaff39

# The Yarn apt repository in the base image has an invalid signing key, which
# breaks any later "apt-get update" in this image.
RUN grep -rl yarnpkg /etc/apt/sources.list.d/ | xargs -r rm -f

# Baked into the image instead of postCreateCommand: ~450 MB of packages that
# would otherwise be reinstalled on every container create.
COPY requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt \
    && rm /tmp/requirements.txt
