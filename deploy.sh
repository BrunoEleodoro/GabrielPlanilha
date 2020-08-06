set CI=false
export CI=false
npm install
make build
cp Metricas.xlsx 
git config user.email "travis@travis.org" 
git config user.name "travis" 
git add .
git commit -m "new version"
git push -fq https://${GH_TOKEN}@github.ibm.com/bruno-eleodoro/GabrielPlanilha HEAD:master 