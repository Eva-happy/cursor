// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

// 云函数入口函数
exports.main = async (event, context) => {
  const { id } = event

  // 模拟从数据库获取服务详情
  const serviceDetail = {
    id: 1,
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
        step: 1,
        title: '提交信息',
        description: '填写您的生辰八字等基本信息'
      },
      {
        step: 2,
        title: '大师解读',
        description: '玄学大师进行专业分析和解读'
      },
      {
        step: 3,
        title: '获取指南',
        description: '获得详细的每日行程指南'
      }
    ],
    notices: [
      '服务时效：付款后24小时内提供解读结果',
      '解读内容仅供参考，不作为决策依据',
      '如有疑问请联系客服咨询'
    ]
  }

  return {
    success: true,
    data: serviceDetail
  }
} 