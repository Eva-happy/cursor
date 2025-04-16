const app = getApp()

Page({
  data: {
    services: [
      {
        id: 1,
        name: '每日行程指南',
        price: 9.9
      },
      {
        id: 2,
        name: '事业运势分析',
        price: 19.9
      },
      {
        id: 3,
        name: '感情运势指导',
        price: 19.9
      },
      {
        id: 4,
        name: '综合运势预测',
        price: 39.9
      }
    ],
    dailyFortune: {
      content: '今日整体运势不错，适合外出谈事，贵人运旺盛。建议把握机会，主动出击。',
      direction: '东南',
      color: '紫色',
      number: '6、8'
    },
    consultants: [
      {
        id: 1,
        name: '张大师',
        title: '玄学专家',
        rating: 4.9,
        consultCount: 1280
      },
      {
        id: 2,
        name: '李大师',
        title: '命理大师',
        rating: 4.8,
        consultCount: 960
      },
      {
        id: 3,
        name: '王大师',
        title: '风水专家',
        rating: 4.9,
        consultCount: 1560
      }
    ],
    reviews: [
      {
        id: 1,
        userName: '张先生',
        rating: 5.0,
        date: '2024-01-28',
        content: '大师的建议非常准确，按照指引确实遇到了贵人相助，事业有了新的突破。'
      },
      {
        id: 2,
        userName: '李女士',
        rating: 4.9,
        date: '2024-01-27',
        content: '感谢大师的指点，让我在感情方面豁然开朗，现在已经遇到了理想的对象。'
      }
    ]
  },

  onLoad: function () {
    this.getDailyFortune()
  },

  // 获取每日运势
  getDailyFortune: function () {
    // 这里可以调用云函数获取每日运势数据
    wx.cloud.callFunction({
      name: 'getDailyFortune',
      success: res => {
        if (res.result) {
          this.setData({
            dailyFortune: res.result
          })
        }
      },
      fail: err => {
        console.error('获取每日运势失败：', err)
      }
    })
  },

  // 跳转到服务详情页
  navigateToService: function (e) {
    const serviceId = e.currentTarget.dataset.id
    wx.navigateTo({
      url: `/pages/service/service?id=${serviceId}`
    })
  },

  // 显示咨询师详情
  showConsultantDetail: function (e) {
    const consultantId = e.currentTarget.dataset.id
    wx.navigateTo({
      url: `/pages/consultant/consultant?id=${consultantId}`
    })
  },

  // 分享功能
  onShareAppMessage: function () {
    return {
      title: '玄学人生指南 - 您的专属运势顾问',
      path: '/pages/index/index'
    }
  }
}) 