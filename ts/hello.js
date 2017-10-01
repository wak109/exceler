function wait() {
    return new Promise(function (resolve) {
        setTimeout(function () {
            console.log('wait');
        }, 2000);
    });
}
function main() {
    return wait()
        .then(function () {
        return wait();
    })
        .then(function () {
        return wait();
    });
}
