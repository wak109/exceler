
function wait(): Promise<any> {
      return new Promise(resolve => {
              setTimeout(() => {
                        console.log('wait');
                      }, 2000);
            })
}

function main() {
      return wait()
        .then(() => {
                  return wait();
                })
        .then(() => {
                  return wait();
                });
}

