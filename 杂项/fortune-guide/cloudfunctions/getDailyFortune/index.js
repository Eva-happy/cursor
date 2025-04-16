// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const wxContext = cloud.getWXContext()
  const userId = wxContext.OPENID
  
  try {
    // 获取用户的生辰八字信息
    const userInfo = await db.collection('users').doc(userId).get()
    
    // 根据日期生成运势
    const today = new Date()
    const dateStr = `${today.getFullYear()}-${today.getMonth() + 1}-${today.getDate()}`
    
    // 使用简单算法生成每日运势
    const directions = ['东', '南', '西', '北', '东南', '东北', '西南', '西北']
    const colors = ['红色', '橙色', '黄色', '绿色', '蓝色', '紫色', '白色', '黑色']
    const numbers = ['1,6', '2,7', '3,8', '4,9', '5,0']
    
    // 使用日期作为种子生成随机数
    const seed = Date.parse(dateStr)
    const random = (seed) => {
      const x = Math.sin(seed) * 10000
      return x - Math.floor(x)
    }
    
    // 生成今日运势
    const fortune = {
      date: dateStr,
      content: generateFortuneContent(random(seed)),
      direction: directions[Math.floor(random(seed + 1) * directions.length)],
      color: colors[Math.floor(random(seed + 2) * colors.length)],
      number: numbers[Math.floor(random(seed + 3) * numbers.length)]
    }
    
    // 保存到数据库
    await db.collection('fortunes').add({
      data: {
        userId,
        ...fortune,
        createdAt: db.serverDate()
      }
    })
    
    return fortune
    
  } catch (err) {
    console.error(err)
    return {
      error: err.message
    }
  }
}

// 生成运势内容
function generateFortuneContent(random) {
  const contents = [
    '今日运势不错，适合外出谈事，贵人运旺盛。建议把握机会，主动出击。',
    '今日财运亨通，适合投资理财，但需谨慎决策。建议多与贵人沟通交流。',
    '今日桃花运旺，易遇贵人，适合社交活动。建议着重人际关系的经营。',
    '今日事业运佳，工作顺利，易有新的机遇。建议积极进取，把握机会。',
    '今日平稳运势，适合处理日常事务。建议保持平和心态，循序渐进。'
  ]
  
  return contents[Math.floor(random * contents.length)]
} 