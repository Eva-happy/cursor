App({
  globalData: {
    userInfo: null,
    services: [
      {
        id: 1,
        name: '每日行程指南',
        price: 9.9,
        description: '为您分析今日最适宜的行程方向、吉利颜色和数字'
      },
      {
        id: 2,
        name: '事业运势分析',
        price: 19.9,
        description: '深度解读您的事业发展方向，助您把握机遇'
      },
      {
        id: 3,
        name: '感情运势指导',
        price: 19.9,
        description: '专业解读感情运势，助您找到真爱良缘'
      },
      {
        id: 4,
        name: '综合运势预测',
        price: 39.9,
        description: '全方位运势分析，助您趋吉避凶'
      }
    ]
  },
  onLaunch: function () {
    // 初始化云开发
    if (!wx.cloud) {
      console.error('请使用 2.2.3 或以上的基础库以使用云能力')
    } else {
      wx.cloud.init({
        env: 'fortune-guide-cloud', // 替换为你创建的环境ID
        traceUser: true
      })
    }

    // 获取用户信息
    wx.getSetting({
      success: res => {
        if (res.authSetting['scope.userInfo']) {
          wx.getUserInfo({
            success: res => {
              this.globalData.userInfo = res.userInfo
              if (this.userInfoReadyCallback) {
                this.userInfoReadyCallback(res)
              }
            }
          })
        }
      }
    })
  }
}) 