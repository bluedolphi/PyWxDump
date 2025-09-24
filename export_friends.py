# -*- coding: utf-8 -*-
"""
微信好友信息导出工具
基于PyWxDump库导出微信好友详细信息
"""
import json
import os
import csv
from typing import Dict, List, Any
from datetime import datetime

import pywxdump
from pywxdump import MicroHandler, get_wx_info, get_wx_db


class WxFriendsExporter:
    """微信好友信息导出器"""
    
    def __init__(self, db_path: str = None, key: str = None):
        """
        初始化好友导出器
        
        Args:
            db_path: 微信数据库路径
            key: 数据库解密密钥
        """
        self.db_path = db_path
        self.key = key
        self.micro_handler = None
        
    def initialize(self) -> bool:
        """初始化数据库连接"""
        try:
            if not self.db_path or not self.key:
                # 如果没有提供路径和密钥，尝试自动获取
                wx_info = get_wx_info()
                if not wx_info:
                    print("未检测到运行中的微信进程")
                    return False
                
                # 使用第一个微信实例
                wx_instance = wx_info[0]
                wxid = wx_instance.get('wxid')
                self.key = wx_instance.get('key')
                
                # 获取数据库路径
                db_info = get_wx_db(wxid)
                if not db_info:
                    print("无法获取微信数据库路径")
                    return False
                
                # 找到MicroMsg数据库路径
                micromsg_path = None
                for db_item in db_info:
                    if db_item.get('db_type') == 'MicroMsg':
                        micromsg_path = db_item.get('db_path')
                        break
                
                if not micromsg_path:
                    print("无法找到MicroMsg数据库")
                    return False
                
                self.db_path = micromsg_path
                
                if not self.db_path or not self.key:
                    print("无法获取微信数据库路径或密钥")
                    return False
                
                print(f"检测到微信账号: {wxid}")
            
            # 先解密数据库文件
            import tempfile
            import os
            
            # 创建临时文件存放解密后的数据库
            temp_dir = tempfile.gettempdir()
            decrypted_db_path = os.path.join(temp_dir, f"decrypted_MicroMsg_{wxid}.db")
            
            # 使用PyWxDump的解密功能
            decrypt_result = pywxdump.decrypt(self.key, self.db_path, decrypted_db_path)
            if not decrypt_result:
                print("数据库解密失败")
                return False
            
            print(f"数据库已解密到: {decrypted_db_path}")
            
            # 配置数据库连接
            db_config = {
                "key": "decrypted",  # 解密后不需要密钥
                "path": decrypted_db_path,
                "type": "sqlite"
            }
            
            # 初始化MicroMsg数据库处理器
            self.micro_handler = MicroHandler(db_config)
            print("数据库连接初始化成功")
            return True
            
        except Exception as e:
            print(f"初始化失败: {e}")
            return False
    
    def get_friends_info(self) -> List[Dict[str, Any]]:
        """
        获取好友信息列表
        
        Returns:
            好友信息列表
        """
        if not self.micro_handler:
            if not self.initialize():
                return []
        
        try:
            # 查询Contact表获取好友信息
            sql = """
            SELECT 
                UserName,          -- 微信号/用户ID
                Alias,             -- 微信号
                NickName,          -- 昵称
                Remark,            -- 备注
                Type,              -- 联系人类型
                VerifyFlag,        -- 验证标志
                Reserved1,         -- 保留字段1
                Reserved2,         -- 保留字段2
                ExtraBuf           -- 额外信息缓冲区
            FROM Contact 
            WHERE Type & 1 = 1    -- 个人好友类型
            AND UserName NOT LIKE 'gh_%'  -- 排除公众号
            AND DelFlag = 0       -- 未删除的联系人
            ORDER BY NickName COLLATE NOCASE;
            """
            
            result = self.micro_handler.execute(sql)
            
            if not result:
                print("未找到好友信息")
                return []
            
            friends = []
            for row in result:
                (UserName, Alias, NickName, Remark, Type, VerifyFlag, Reserved1, Reserved2, ExtraBuf) = row
                
                friend_info = {
                    'wxid': UserName,
                    'wechat_id': Alias or '',
                    'nickname': NickName or '',
                    'remark': Remark or '',
                    'contact_type': Type,
                    'verify_flag': VerifyFlag,
                    'gender': self._extract_gender(ExtraBuf),
                    'region': self._extract_region(ExtraBuf),
                    'signature': self._extract_signature(ExtraBuf),
                    'avatar_url': self._get_avatar_url(UserName)
                }
                
                friends.append(friend_info)
            
            print(f"找到 {len(friends)} 个好友")
            return friends
            
        except Exception as e:
            print(f"获取好友信息失败: {e}")
            return []
    
    def _extract_gender(self, extra_buf: bytes) -> int:
        """从ExtraBuf中提取性别信息"""
        if not extra_buf:
            return 0
        
        try:
            # 性别信息通常在ExtraBuf的特定位置
            # 0: 未知, 1: 男, 2: 女
            if len(extra_buf) > 8:
                gender_byte = extra_buf[8]
                return gender_byte if gender_byte in [0, 1, 2] else 0
        except:
            pass
        return 0
    
    def _extract_region(self, extra_buf: bytes) -> str:
        """从ExtraBuf中提取地区信息"""
        if not extra_buf:
            return ""
        
        try:
            # 地区信息解析（简化版本）
            import blackboxprotobuf
            if extra_buf:
                decoded = blackboxprotobuf.decode_message(extra_buf)
                # 根据实际protobuf结构解析地区信息
                return str(decoded)[:50]  # 限制长度
        except:
            pass
        return ""
    
    def _extract_signature(self, extra_buf: bytes) -> str:
        """从ExtraBuf中提取个性签名"""
        if not extra_buf:
            return ""
        
        try:
            # 个性签名解析（简化版本）
            import blackboxprotobuf
            if extra_buf:
                decoded = blackboxprotobuf.decode_message(extra_buf)
                # 根据实际protobuf结构解析签名信息
                return str(decoded)[:100]  # 限制长度
        except:
            pass
        return ""
    
    def _get_avatar_url(self, username: str) -> str:
        """获取头像URL"""
        try:
            if self.micro_handler.tables_exist("ContactHeadImgUrl"):
                sql = "SELECT bigHeadImgUrl FROM ContactHeadImgUrl WHERE usrName = ?"
                result = self.micro_handler.execute(sql, (username,))
                if result and result[0][0]:
                    return result[0][0]
        except:
            pass
        return ""
    
    def export_to_json(self, output_path: str = None) -> bool:
        """
        导出好友信息到JSON文件
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            导出是否成功
        """
        friends = self.get_friends_info()
        if not friends:
            return False
        
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"out/wx_friends_{timestamp}.json"
        else:
            # 确保路径在out目录下
            if not output_path.startswith('out/'):
                output_path = f"out/{output_path}"
        
        try:
            export_data = {
                "export_time": datetime.now().isoformat(),
                "total_count": len(friends),
                "friends": friends
            }
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            
            print(f"好友信息已导出到: {output_path}")
            return True
            
        except Exception as e:
            print(f"JSON导出失败: {e}")
            return False
    
    def export_to_csv(self, output_path: str = None) -> bool:
        """
        导出好友信息到CSV文件
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            导出是否成功
        """
        friends = self.get_friends_info()
        if not friends:
            return False
        
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"out/wx_friends_{timestamp}.csv"
        else:
            # 确保路径在out目录下
            if not output_path.startswith('out/'):
                output_path = f"out/{output_path}"
        
        try:
            with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                
                # 写入表头
                headers = [
                    '微信ID', '微信号', '昵称', '备注', 
                    '性别', '地区', '个性签名', '联系人类型'
                ]
                writer.writerow(headers)
                
                # 写入好友数据
                gender_map = {0: '未知', 1: '男', 2: '女'}
                
                for friend in friends:
                    writer.writerow([
                        friend['wxid'],
                        friend['wechat_id'],
                        friend['nickname'],
                        friend['remark'],
                        gender_map.get(friend['gender'], '未知'),
                        friend['region'],
                        friend['signature'],
                        friend['contact_type']
                    ])
            
            print(f"好友信息已导出到: {output_path}")
            return True
            
        except Exception as e:
            print(f"CSV导出失败: {e}")
            return False
    
    def print_friends_summary(self):
        """打印好友信息摘要"""
        friends = self.get_friends_info()
        if not friends:
            return
        
        print("\n" + "="*60)
        print("微信好友信息摘要")
        print("="*60)
        print(f"好友总数: {len(friends)}")
        
        # 统计性别
        gender_count = {0: 0, 1: 0, 2: 0}
        for friend in friends:
            gender_count[friend['gender']] += 1
        
        print(f"男性好友: {gender_count[1]}")
        print(f"女性好友: {gender_count[2]}")
        print(f"未知性别: {gender_count[0]}")
        
        # 显示前20个好友
        print(f"\n好友列表 (显示前20个):")
        print("-" * 80)
        for i, friend in enumerate(friends[:20]):
            nickname = friend['nickname'][:15] if friend['nickname'] else '无昵称'
            remark = friend['remark'][:15] if friend['remark'] else '无备注'
            wechat_id = friend['wechat_id'][:15] if friend['wechat_id'] else '未设置'
            
            print(f"{i+1:2d}. {nickname:15s} | {remark:15s} | {wechat_id}")
        
        if len(friends) > 20:
            print(f"... 还有 {len(friends) - 20} 个好友")
        
        print("="*60)


def main():
    """主函数"""
    print("微信好友信息导出工具")
    print("=" * 40)
    
    try:
        # 创建导出器实例
        exporter = WxFriendsExporter()
        
        # 打印好友摘要
        exporter.print_friends_summary()
        
        # 导出到不同格式
        print("\n正在导出文件...")
        
        # 导出到JSON
        if exporter.export_to_json():
            print("JSON导出成功")
        
        # 导出到CSV
        if exporter.export_to_csv():
            print("CSV导出成功")
        
        print("\n好友信息导出完成!")
        
    except Exception as e:
        print(f"程序执行失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()