module.exports = {
  apps: [{
    name: 'ims-preview',
    script: 'python3',
    args: 'app.py',
    cwd: '/home/user/webapp',
    watch: false,
    instances: 1,
    exec_mode: 'fork',
    error_file: '/home/user/webapp/logs/err.log',
    out_file: '/home/user/webapp/logs/out.log',
    env: {
      PORT: '3000',
      DOCS_ROOT: '/home/user/webapp/docs'
    }
  }]
}
