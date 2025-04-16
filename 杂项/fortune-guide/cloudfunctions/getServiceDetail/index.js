// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const { serviceId } = event
  const wxContext = cloud.getWXContext()
  
  try {
    // 从数据库获取服务详情
    const serviceData = await db.collection('services').doc(serviceId).get()
    
    // 如果没有找到服务，返回默认数据
    if (!serviceData.data) {
      return {
        id: serviceId,
        name: '每日行程指南',
        price: 9.9,
        coverImage: '/images/service-cover.jpg',
        description: '通过专业的玄学分析，为您提供每日最适宜的行程方向、吉利颜色和数字，助您趋吉避凶，把握机遇。',
        features: [
          {
            id: 1,
            icon: '/images/feature1.png',
            title: '专业解读',
            description: '资深玄学大师提供专业解读'
          },
          {
            id: 2,
            icon: '/images/feature2.png',
            title: '个性化指导',
            description: '根据您的生辰八字提供个性化指导'
          },
          {
            id: 3,
            icon: '/images/feature3.png',
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
    }
    
    return serviceData.data
    
  } catch (err) {
    console.error(err)
    return {
      error: err.message
    }
  }
} 