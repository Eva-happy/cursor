const app = getApp()

Page({
  data: {
    service: {
      id: null,
      name: '',
      price: 0,
      coverImage: '',
      description: '',
      features: [],
      process: [],
      notices: []
    },
    showPayment: false,
    selectedPayment: 'wxpay'
  },

  onLoad: function (options) {
    const serviceId = options.id
    this.getServiceDetail(serviceId)
  },

  // 获取服务详情
  getServiceDetail: function (serviceId) {
    // 调用云函数获取服务详情
    wx.cloud.callFunction({
      name: 'getServiceDetail',
      data: {
        serviceId: serviceId
      },
      success: res => {
        if (res.result) {
          this.setData({
            service: {
              id: serviceId,
              name: '每日行程指南',
              price: 9.9,
              coverImage: '',
              description: '通过专业的玄学分析，为您提供每日最适宜的行程方向、吉利颜色和数字，助您趋吉避凶，把握机遇。',
              features: [
                {
                  id: 1,
                  icon: '',
                  title: '专业解读',
                  description: '资深玄学大师提供专业解读'
                },
                {
                  id: 2,
                  icon: '',
                  title: '个性化指导',
                  description: '根据您的生辰八字提供个性化指导'
                },
                {
                  id: 3,
                  icon: '',
                  title: '实时更新',
                  description: '每日运势实时更新'
                }
              ],
              process: [
                {
                  step: '1',
                  title: '提交信息',
                  description: '填写您的生辰八字等基本信息'
                },
                {
                  step: '2',
                  title: '大师解读',
                  description: '玄学大师进行专业分析和解读'
                },
                {
                  step: '3',
                  title: '获取指引',
                  description: '获得详细的每日行程指引'
                }
              ],
              notices: [
                '服务购买后即时生效',
                '每日运势于凌晨0点更新',
                '本服务仅供参考，不作为决策依据',
                '如有疑问请联系客服咨询'
              ]
            }
          })
        }
      },
      fail: err => {
        console.error('获取服务详情失败：', err)
        wx.showToast({
          title: '获取服务详情失败',
          icon: 'none'
        })
      }
    })
  },

  // 联系客服
  contactService: function () {
    wx.showToast({
      title: '正在接入客服...',
      icon: 'loading',
      duration: 2000
    })
  },

  // 显示支付弹窗
  handlePurchase: function () {
    this.setData({
      showPayment: true
    })
  },

  // 关闭支付弹窗
  closePayment: function () {
    this.setData({
      showPayment: false
    })
  },

  // 选择支付方式
  selectPayment: function (e) {
    const method = e.currentTarget.dataset.method
    this.setData({
      selectedPayment: method
    })
  },

  // 确认支付
  confirmPayment: function () {
    const that = this
    // 调用支付接口
    wx.cloud.callFunction({
      name: 'createOrder',
      data: {
        serviceId: that.data.service.id,
        amount: that.data.service.price
      },
      success: res => {
        if (res.result && res.result.payment) {
          // 调用微信支付
          wx.requestPayment({
            ...res.result.payment,
            success: function () {
              wx.showToast({
                title: '支付成功',
                icon: 'success'
              })
              that.setData({
                showPayment: false
              })
              // 跳转到订单页面
              wx.navigateTo({
                url: '/pages/order/order'
              })
            },
            fail: function (err) {
              console.error('支付失败：', err)
              wx.showToast({
                title: '支付失败',
                icon: 'none'
              })
            }
          })
        }
      },
      fail: err => {
        console.error('创建订单失败：', err)
        wx.showToast({
          title: '创建订单失败',
          icon: 'none'
        })
      }
    })
  },

  // 分享功能
  onShareAppMessage: function () {
    return {
      title: this.data.service.name,
      path: `/pages/service/service?id=${this.data.service.id}`,
      imageUrl: ''
    }
  }
}) 