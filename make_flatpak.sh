export LC_ALL=en_US.utf8
sudo apt install docker.io docker-buildx -y
python3 -m pip install briefcase
rm -rf build/
rm -rf dist/
rm -rf logs/
python3 -m briefcase create linux flatpak
python3 -m briefcase build linux flatpak
python3 -m briefcase package linux flatpak
