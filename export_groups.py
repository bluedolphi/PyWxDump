# -*- coding: utf-8 -*-
"""
微信群信息导出工具
基于PyWxDump库导出微信群详细信息
"""
import json
import os
import csv
from typing import Dict, List, Any
from datetime import datetime

import pywxdump
from pywxdump import MicroHandler, get_wx_info, get_wx_db


class WXGroupsExporter:
    """微信群信息导出器"""
    
    def __init__(self, db_path: str = None, key: str = None):
        """
        初始化群导出器
        
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
    
    def get_groups_info(self) -> List[Dict[str, Any]]:
        """
        获取群信息列表
        
        Returns:
            群信息列表
        """
        if not self.micro_handler:
            if not self.initialize():
                return []
        
        try:
            # 查询ChatRoom表获取群基本信息
            chatroom_sql = """
            SELECT 
                ChatRoomName,     -- 群ID
                SelfDisplayName,  -- 自己在群中的显示名称
                Owner,            -- 群主
                ChatRoomFlag,     -- 群标志
                UserNameList,     -- 用户名列表
                DisplayNameList,  -- 显示名称列表
                RoomData          -- 房间数据
            FROM ChatRoom 
            WHERE ChatRoomName IS NOT NULL
            ORDER BY ChatRoomName;
            """
            
            chatroom_result = self.micro_handler.execute(chatroom_sql)
            
            groups = []
            
            if chatroom_result:
                for row in chatroom_result:
                    (ChatRoomName, SelfDisplayName, Owner, ChatRoomFlag, 
                     UserNameList, DisplayNameList, RoomData) = row
                    
                    # 从Contact表获取群名
                    group_name = self._get_group_name_from_contact(ChatRoomName)
                    
                    # 计算成员数量 - 从UserNameList中提取实际成员
                    member_count = self._calculate_actual_members(UserNameList)
                    
                    group_info = {
                        'group_id': ChatRoomName or '',
                        'display_name': group_name or SelfDisplayName or '未命名群',
                        'member_count': member_count,
                        'owner_id': Owner or 0,
                        'chatroom_flag': ChatRoomFlag or 0,
                        'self_display_name': SelfDisplayName or '',
                        'user_name_list': UserNameList or '',
                        'display_name_list': DisplayNameList or '',
                        'room_data': str(RoomData) if RoomData else ''
                    }
                    
                    groups.append(group_info)
            
            print(f"找到 {len(groups)} 个群聊")
            return groups
            
        except Exception as e:
            print(f"获取群信息失败: {e}")
            return []
    
    def _get_group_name_from_contact(self, chatroom_name: str) -> str:
        """从Contact表获取群名"""
        try:
            if self.micro_handler.tables_exist("Contact"):
                sql = "SELECT NickName, Remark FROM Contact WHERE UserName = ? AND Type = 3"
                result = self.micro_handler.execute(sql, (chatroom_name,))
                if result and result[0]:
                    nick_name, remark = result[0]
                    # 优先使用备注，其次使用昵称
                    return remark or nick_name or ''
        except Exception as e:
            print(f"获取群名失败: {e}")
        return ''
    
    def _calculate_actual_members(self, username_list: str) -> int:
        """计算实际的群成员数量"""
        if not username_list:
            return 0
        
        try:
            # UserNameList格式通常是 ^wxid1^wxid2^... 或类似格式
            # 去除首尾的分隔符并分割
            members = username_list.strip('^').split('^')
            # 过滤空字符串和无效项
            actual_members = [m for m in members if m and m.strip()]
            return len(actual_members)
        except:
            # 如果解析失败，使用原来的方法
            return len(username_list.split('^')) if username_list else 0
    
    def _extract_owner(self, extra_buf: bytes) -> str:
        """从ExtraBuf中提取群主信息"""
        if not extra_buf:
            return ""
        
        try:
            import blackboxprotobuf
            decoded = blackboxprotobuf.decode_message(extra_buf)
            # 根据实际protobuf结构解析群主信息
            return str(decoded)[:50]  # 简化版本
        except:
            pass
        return ""
    
    def _extract_notice(self, extra_buf: bytes) -> str:
        """从ExtraBuf中提取群公告"""
        if not extra_buf:
            return ""
        
        try:
            import blackboxprotobuf
            decoded = blackboxprotobuf.decode_message(extra_buf)
            # 根据实际protobuf结构解析群公告
            return str(decoded)[:200]  # 简化版本，限制长度
        except:
            pass
        return ""
    
    def _get_group_members(self, chatroom_name: str) -> List[Dict[str, Any]]:
        """获取群成员信息"""
        try:
            if not self.micro_handler.tables_exist("ChatRoomInfo"):
                return []
            
            sql = """
            SELECT 
                MemberList,       -- 成员列表
                DisplayName,      -- 显示名称
                ChatRoomName      -- 群名称
            FROM ChatRoomInfo 
            WHERE ChatRoomName = ?;
            """
            
            result = self.micro_handler.execute(sql, (chatroom_name,))
            
            if not result:
                return []
            
            members = []
            for row in result:
                (MemberList, DisplayName, ChatRoomName) = row
                
                if MemberList:
                    # MemberList通常是protobuf编码的成员列表
                    member_ids = self._parse_member_list(MemberList)
                    for member_id in member_ids:
                        member_info = {
                            'wxid': member_id,
                            'display_name': DisplayName or '',
                            'role': self._get_member_role(member_id, chatroom_name)
                        }
                        members.append(member_info)
            
            return members
            
        except Exception as e:
            print(f"❌ 获取群成员失败: {e}")
            return []
    
    def _parse_member_list(self, member_list_bytes: bytes) -> List[str]:
        """解析成员列表字节码"""
        try:
            import blackboxprotobuf
            decoded = blackboxprotobuf.decode_message(member_list_bytes)
            # 根据实际protobuf结构解析成员ID列表
            # 这里返回简化的版本，实际需要根据protobuf结构解析
            if isinstance(decoded, dict):
                for key, value in decoded.items():
                    if isinstance(value, list):
                        return [str(v) for v in value if v]
            return []
        except:
            return []
    
    def _get_member_role(self, member_id: str, chatroom_name: str) -> str:
        """获取成员角色（群主、管理员、普通成员）"""
        # 这里需要根据实际数据库结构查询成员角色
        # 目前返回简化版本
        return "member"
    
    def get_group_statistics(self) -> Dict[str, Any]:
        """获取群统计信息"""
        groups = self.get_groups_info()
        if not groups:
            return {}
        
        stats = {
            'total_groups': len(groups),
            'total_members': sum(g['member_count'] for g in groups),
            'avg_members_per_group': sum(g['member_count'] for g in groups) / len(groups) if groups else 0,
            'largest_group': max(groups, key=lambda x: x['member_count']) if groups else None,
            'smallest_group': min(groups, key=lambda x: x['member_count']) if groups else None,
            'groups_by_size': {}
        }
        
        # 按规模分类
        for group in groups:
            size = group['member_count']
            if size <= 10:
                category = '小群(1-10人)'
            elif size <= 50:
                category = '中群(11-50人)'
            elif size <= 100:
                category = '大群(51-100人)'
            else:
                category = '超大群(100+人)'
            
            stats['groups_by_size'][category] = stats['groups_by_size'].get(category, 0) + 1
        
        return stats
    
    def export_to_json(self, output_path: str = None) -> bool:
        """
        导出群信息到JSON文件
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            导出是否成功
        """
        groups = self.get_groups_info()
        if not groups:
            return False
        
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"out/wx_groups_{timestamp}.json"
        else:
            # 确保路径在out目录下
            if not output_path.startswith('out/'):
                output_path = f"out/{output_path}"
        
        try:
            export_data = {
                "export_time": datetime.now().isoformat(),
                "statistics": self.get_group_statistics(),
                "total_count": len(groups),
                "groups": groups
            }
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            
            print(f"群信息已导出到: {output_path}")
            return True
            
        except Exception as e:
            print(f"JSON导出失败: {e}")
            return False
    
    def export_to_csv(self, output_path: str = None) -> bool:
        """
        导出群信息到CSV文件
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            导出是否成功
        """
        groups = self.get_groups_info()
        if not groups:
            return False
        
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"out/wx_groups_{timestamp}.csv"
        else:
            # 确保路径在out目录下
            if not output_path.startswith('out/'):
                output_path = f"out/{output_path}"
        
        try:
            with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                
                # 写入表头
                headers = [
                    '群ID', '群名称', '成员数量', '群主ID', 
                    '群标志', '自己在群中的昵称', '用户名列表', '显示名称列表'
                ]
                writer.writerow(headers)
                
                # 写入群数据
                for group in groups:
                    writer.writerow([
                        group['group_id'],
                        group['display_name'],
                        group['member_count'],
                        group['owner_id'],
                        group['chatroom_flag'],
                        group['self_display_name'],
                        group['user_name_list'][:100] if group['user_name_list'] else '',  # 限制长度
                        group['display_name_list'][:100] if group['display_name_list'] else ''  # 限制长度
                    ])
            
            print(f"群信息已导出到: {output_path}")
            return True
            
        except Exception as e:
            print(f"CSV导出失败: {e}")
            return False
    
    def print_groups_summary(self):
        """打印群信息摘要"""
        groups = self.get_groups_info()
        if not groups:
            return
        
        stats = self.get_group_statistics()
        
        print("\n" + "="*60)
        print("微信群信息摘要")
        print("="*60)
        print(f"群总数: {stats['total_groups']}")
        print(f"总成员数: {stats['total_members']}")
        print(f"平均群成员数: {stats['avg_members_per_group']:.1f}")
        
        if stats['largest_group']:
            print(f"最大群: {stats['largest_group']['display_name']} ({stats['largest_group']['member_count']}人)")
        
        if stats['smallest_group']:
            print(f"最小群: {stats['smallest_group']['display_name']} ({stats['smallest_group']['member_count']}人)")
        
        print("\n群规模分布:")
        for category, count in stats['groups_by_size'].items():
            print(f"  - {category}: {count}个")
        
        # 显示前10个群
        print(f"\n群列表 (显示前10个):")
        print("-" * 80)
        for i, group in enumerate(groups[:10]):
            name = group['display_name'][:20] if group['display_name'] else '未命名群'
            # 过滤特殊字符
            try:
                name = name.encode('gbk', 'ignore').decode('gbk')
            except:
                name = '未命名群'
            members = group['member_count']
            owner = str(group['owner_id'])[:15] if group['owner_id'] else '未知'
            
            print(f"{i+1:2d}. {name:20s} | {members:3d}人 | 群主: {owner}")
        
        if len(groups) > 10:
            print(f"... 还有 {len(groups) - 10} 个群")
        
        print("="*60)


def main():
    """主函数"""
    print("微信群信息导出工具")
    print("=" * 40)
    
    try:
        # 创建导出器实例
        exporter = WXGroupsExporter()
        
        # 打印群摘要
        exporter.print_groups_summary()
        
        # 导出到不同格式
        print("\n正在导出文件...")
        
        # 导出到JSON
        if exporter.export_to_json():
            print("JSON导出成功")
        
        # 导出到CSV
        if exporter.export_to_csv():
            print("CSV导出成功")
        
        print("\n群信息导出完成!")
        
    except Exception as e:
        print(f"程序执行失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()